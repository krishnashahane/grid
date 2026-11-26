import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  base: './',
  build: {
    rolldownOptions: {
      output: {
        // Framework and icon code changes rarely, so it gets its own long-cached chunk.
        codeSplitting: {
          groups: [
            { name: 'vendor', test: /node_modules[\\/](react|react-dom|scheduler|lucide-react)[\\/]/ },
          ],
        },
      },
    },
  },
  test: { globals: true, environment: 'node', include: ['tests/**/*.test.ts'] },
} as any);
