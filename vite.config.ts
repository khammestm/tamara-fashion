import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import {fileURLToPath} from 'url';
import {defineConfig} from 'vite';

const projectRoot = path.dirname(fileURLToPath(import.meta.url));

// Copy the root-level images/ folder into dist/ so index.html can read
// ./images/* after `vite build` (public/ is copied automatically, images/ is not).
const copyImagesPlugin = () => ({
  name: 'copy-images-to-dist',
  apply: 'build' as const,
  closeBundle() {
    const src = path.resolve(projectRoot, 'images');
    const dest = path.resolve(projectRoot, 'dist', 'images');
    if (fs.existsSync(src)) {
      fs.mkdirSync(dest, {recursive: true});
      fs.cpSync(src, dest, {recursive: true});
    }
  },
});

export default defineConfig(() => {
  return {
    plugins: [react(), tailwindcss(), copyImagesPlugin()],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modify—file watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
      // Disable file watching when DISABLE_HMR is true to save CPU during agent edits.
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
  };
});
