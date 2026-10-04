import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { helperSourcePath } from '../src/helper-client.js'
import { name } from '../src/index.js'
import { DEFAULTS, resolveConfig } from '../src/settings.js'

const pkg = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')) as { name: string; files: string[]; publishConfig?: { access?: string } }

describe('packaging', () => {
  it('publishes publicly under the plugin name', () => {
    expect(pkg.name).toBe(name)
    expect(pkg.publishConfig?.access).toBe('public')
    expect(readFileSync(new URL('../cordis.patch.yml', import.meta.url), 'utf8')).toContain(pkg.name)
  })
  it('ships the helper source that the runtime compiles', () => {
    expect(pkg.files).toContain('helper/OfficeHelper.cs')
    expect(existsSync(helperSourcePath())).toBe(true)
  })
  it('keeps the helper within C# 5 (in-box csc)', () => {
    const source = readFileSync(helperSourcePath(), 'utf8')
    expect(source).not.toMatch(/\$"/)
    expect(source).not.toMatch(/\?\.[A-Za-z]/)
  })
})

const clientBundle = new URL('../lib/client.js', import.meta.url)

describe('client bundle', () => {
  it.skipIf(!existsSync(clientBundle))('registers under the package name', () => {
    expect(readFileSync(clientBundle, 'utf8').slice(0, 200)).toContain(`id: "${pkg.name}"`)
  })
})

describe('settings', () => {
  it('fills defaults, unwraps volatile refs and clamps', () => {
    expect(resolveConfig({})).toEqual(DEFAULTS)
    expect(resolveConfig({ renderWidth: { get: () => 1400 }, showOnOpen: false })).toEqual({ renderWidth: 1400, showOnOpen: false })
    expect(resolveConfig({ renderWidth: 99999 }).renderWidth).toBe(2000)
    expect(resolveConfig({ renderWidth: 'wide' }).renderWidth).toBe(DEFAULTS.renderWidth)
  })
})
