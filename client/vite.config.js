import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      '/dart': 'http://localhost:3000'  // 개발 시 백엔드 프록시
    }
  }
});
