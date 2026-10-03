import { defineConfig } from 'astro/config';
import node from '@astrojs/node';

export default defineConfig({
  site: process.env.PUBLIC_SITE_URL || 'https://api.cystems.ec',
  output: 'server',
  adapter: node({ mode: 'standalone' }),
  security: { checkOrigin: true },
  server: { host: true, port: 4321 },
});
