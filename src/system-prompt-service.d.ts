/**
 * Ambient typing for the host's `systemPrompt` service
 * (@deepseek-ai/dsh-system-prompt), mirrored only for the surface this plugin
 * uses so it does not become a hard dependency.
 */
export {}

declare module '@deepseek-ai/cordis' {
  interface Context {
    systemPrompt: {
      /** Register an ordered system-prompt section; returns its disposer. */
      section(section: {
        readonly name: string
        readonly order: number
        readonly text: string | ((context: unknown) => string)
      }): () => void
    }
  }
}
