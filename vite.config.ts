import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { readFileSync } from 'node:fs';

// `vite preview` serves the same security headers as production (single source: vercel.json),
// so the Content Security Policy is tested locally before every deploy.
const deployed = JSON.parse(readFileSync(new URL('./vercel.json', import.meta.url), 'utf8'));
const securityHeaders: Record<string, string> = Object.fromEntries(
  deployed.headers[0].headers.map((h: { key: string; value: string }) => [h.key, h.value]),
);

export default defineConfig({
  plugins: [react()],
  base: './',
  preview: { headers: securityHeaders },
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
