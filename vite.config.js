import { defineConfig } from 'vite';

// The renderer is plain ES modules; Vite is here for the dev server and to
// bundle the one npm dependency (docx). `base: './'` matters: the packaged app
// loads dist/index.html over file://, where absolute asset paths do not resolve.
export default defineConfig({
  root: 'src',
  base: './',
  build: {
    outDir: '../dist',
    emptyOutDir: true,
    target: 'chrome120',
  },
  server: {
    port: 5183,
    strictPort: true,
  },
});
