/** Standalone host bundle output. */
import type { UserConfig } from 'tsdown'

const host: UserConfig = {
  name: '@copylee/dsh-office',
  entry: ['lib/types/index.js'],
  outDir: 'lib',
  format: ['esm'],
  platform: 'node',
  target: 'es2024',
  fixedExtension: false,
  dts: false,
  clean: false,
  // DSH host libraries are inlined on purpose: the plugin must not depend on
  // the exact rc the host ships.
  deps: { onlyBundle: false },
}

/** Browser half: one ModuleLoader-wrapped file; React comes from the host. */
const client: UserConfig = {
  name: '@copylee/dsh-office/client',
  entry: { client: 'lib/types/client/index.js' },
  outDir: 'lib',
  format: 'cjs',
  platform: 'browser',
  target: 'es2022',
  dts: false,
  clean: false,
  deps: { neverBundle: ['react', 'react/jsx-runtime', 'react-dom'], onlyBundle: false },
  // The host webview has no `process` global; bake NODE_ENV at build time.
  define: { 'process.env.NODE_ENV': JSON.stringify('production') },
  outputOptions: {
    entryFileNames: 'client.js',
    banner: 'window.__ModuleLoader__.load({ id: "@copylee/dsh-office", factory: (require) => {',
    footer: 'return module.exports; } });',
    intro: 'var module = { exports: {} }; var exports = module.exports;',
  },
}

export default [host, client]
