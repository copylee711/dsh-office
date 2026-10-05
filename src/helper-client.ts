/**
 * Builds and drives the Windows helper (helper/OfficeHelper.cs): compiled once
 * with the in-box .NET Framework csc.exe into a per-source-hash cache folder,
 * then kept running and spoken to as JSON lines over stdio.
 */
import { spawn, execFile, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync } from 'node:fs'
import { mkdir, readFile, rename, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

export interface HelperEvent { event: string; reason?: string }

export interface HelperLike {
  call<T = unknown>(cmd: string, args?: Record<string, unknown>, timeoutMs?: number): Promise<T>
  onEvent(listener: (event: HelperEvent) => void): () => void
  dispose(): void
}

const REFERENCES = [
  'System.dll', 'System.Core.dll', 'Microsoft.CSharp.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll',
  'System.Web.Extensions.dll',
]

/** A file in helper/ at the package root: one level up from lib/ or src/, two from lib/types/. */
export function helperAssetPath(name: string): string {
  const candidates = [`../helper/${name}`, `../../helper/${name}`].map(rel => fileURLToPath(new URL(rel, import.meta.url)))
  return candidates.find(path => existsSync(path)) ?? candidates[0]!
}

export function helperSourcePath(): string {
  return helperAssetPath('OfficeHelper.cs')
}

function frameworkDir(): string {
  const windir = process.env.SystemRoot ?? process.env.windir ?? 'C:\\Windows'
  for (const arch of ['Framework64', 'Framework']) {
    const dir = join(windir, 'Microsoft.NET', arch, 'v4.0.30319')
    if (existsSync(join(dir, 'csc.exe'))) return dir
  }
  throw new Error('The .NET Framework 4 C# compiler (csc.exe) was not found. It ships with Windows 10/11; enable ".NET Framework 4.8" in Windows Features.')
}

function cacheRoot(): string {
  return join(process.env.LOCALAPPDATA ?? tmpdir(), 'dsh-office')
}

let building: Promise<string> | undefined

/** Compile the helper if this source version has no cached exe yet; returns the exe path. */
export async function ensureHelperExe(source = helperSourcePath()): Promise<string> {
  building ??= (async () => {
    const code = await readFile(source)
    const hash = createHash('sha256').update(code).digest('hex').slice(0, 12)
    // With DSH_OFFICE_DEBUG=1 the helper is built with line numbers for the stack traces it then writes to stderr.
    const debug = process.env.DSH_OFFICE_DEBUG === '1'
    const dir = join(cacheRoot(), `helper-${hash}${debug ? '-debug' : ''}`)
    const exe = join(dir, 'office-helper.exe')
    if (existsSync(exe)) return exe
    await mkdir(dir, { recursive: true })
    const fw = frameworkDir()
    const temp = join(dir, `build-${process.pid}-${Date.now()}.exe`)
    const args = [
      '-nologo', ...(debug ? ['-debug:pdbonly'] : ['-optimize+']), '-target:exe', `-out:${temp}`,
      ...REFERENCES.map(ref => `-r:${ref}`), source,
    ]
    await new Promise<void>((resolve, reject) => {
      execFile(join(fw, 'csc.exe'), args, { windowsHide: true, timeout: 120_000 }, (error, stdout) => {
        if (error) reject(new Error(`Compiling the office helper failed: ${String(stdout).trim() || error.message}`))
        else resolve()
      })
    })
    try {
      await rename(temp, exe)
    } catch (error) {
      // Another host process won the race; its exe is identical.
      await rm(temp, { force: true })
      if (!existsSync(exe)) throw error
    }
    return exe
  })()
  try {
    return await building
  } catch (error) {
    building = undefined
    throw error
  }
}

interface Pending { resolve(value: unknown): void; reject(error: Error): void; timer: NodeJS.Timeout }

export class HelperClient implements HelperLike {
  private child: ChildProcessWithoutNullStreams | undefined
  private starting: Promise<ChildProcessWithoutNullStreams> | undefined
  private readonly pending = new Map<number, Pending>()
  private readonly listeners = new Set<(event: HelperEvent) => void>()
  private nextId = 0
  private disposed = false

  constructor(private readonly log: (message: string) => void = () => {}) {}

  onEvent(listener: (event: HelperEvent) => void): () => void {
    this.listeners.add(listener)
    return () => { this.listeners.delete(listener) }
  }

  async call<T = unknown>(cmd: string, args: Record<string, unknown> = {}, timeoutMs = 20_000): Promise<T> {
    if (this.disposed) throw new Error('office helper is shut down')
    const child = await this.ensure()
    const id = ++this.nextId
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id)
        reject(new Error(`office helper timed out on "${cmd}"`))
        // The helper handles one command at a time; a stuck one blocks the rest.
        this.kill()
      }, timeoutMs)
      this.pending.set(id, { resolve: resolve as (value: unknown) => void, reject, timer })
      child.stdin.write(`${JSON.stringify({ ...args, id, cmd })}\n`)
    })
  }

  dispose(): void {
    this.disposed = true
    this.kill()
  }

  private kill(): void {
    const child = this.child
    this.child = undefined
    this.starting = undefined
    if (child && child.exitCode === null) {
      try { child.stdin.end() } catch { /* already closed */ }
      setTimeout(() => { if (child.exitCode === null) child.kill() }, 800).unref()
    }
  }

  private ensure(): Promise<ChildProcessWithoutNullStreams> {
    if (this.child && this.child.exitCode === null) return Promise.resolve(this.child)
    this.starting ??= this.start().catch((error: unknown) => {
      this.starting = undefined
      throw error
    })
    return this.starting
  }

  private async start(): Promise<ChildProcessWithoutNullStreams> {
    const exe = await ensureHelperExe()
    const child = spawn(exe, [], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true })
    let buffer = ''
    let ready: (() => void) | undefined
    const readyPromise = new Promise<void>((resolve, reject) => {
      ready = resolve
      child.once('error', reject)
      setTimeout(() => reject(new Error('office helper did not start')), 15_000).unref()
    })
    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      buffer += chunk
      let index: number
      while ((index = buffer.indexOf('\n')) >= 0) {
        const line = buffer.slice(0, index).trim()
        buffer = buffer.slice(index + 1)
        if (line) this.handle(line, () => ready?.())
      }
    })
    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk: string) => this.log(`helper stderr: ${chunk.trim()}`))
    child.on('exit', code => {
      if (this.child === child) this.child = undefined
      this.starting = undefined
      for (const [id, entry] of this.pending) {
        clearTimeout(entry.timer)
        entry.reject(new Error(`office helper exited (code ${String(code)})`))
        this.pending.delete(id)
      }
    })
    await readyPromise
    this.child = child
    return child
  }

  private handle(line: string, onReady: () => void): void {
    let message: { id?: number | null; ok?: boolean; result?: unknown; error?: string; code?: string; event?: string; reason?: string }
    try {
      message = JSON.parse(line) as typeof message
    } catch {
      this.log(`helper sent a non-JSON line: ${line.slice(0, 200)}`)
      return
    }
    if (typeof message.event === 'string') {
      if (message.event === 'ready') { onReady(); return }
      for (const listener of this.listeners) {
        try { listener({ event: message.event, ...(message.reason === undefined ? {} : { reason: message.reason }) }) } catch { /* listener bug */ }
      }
      return
    }
    if (typeof message.id !== 'number') return
    const entry = this.pending.get(message.id)
    if (!entry) return
    this.pending.delete(message.id)
    clearTimeout(entry.timer)
    if (message.ok) entry.resolve(message.result ?? null)
    else entry.reject(Object.assign(new Error(message.error ?? 'helper error'), { code: message.code ?? 'ERROR' }))
  }
}
