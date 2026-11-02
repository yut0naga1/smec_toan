import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';

export default defineConfig({
  plugins: [vue()],
  // GitHub Pages（/smec_toan/）でもローカルでも動くよう相対パスで出力する
  base: './',
});
