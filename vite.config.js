import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  plugins: [
    vue(),
    VitePWA({
      registerType: "prompt",
      includeAssets: ["icon-192.png", "icon-512.png"],
      manifest: {
        name: "电子抄表本",
        short_name: "抄表本",
        lang: "zh-CN",
        start_url: "./",
        scope: "./",
        display: "standalone",
        theme_color: "#087f5b",
        background_color: "#ffffff",
        icons: [192, 512].map((size) => ({
          src: `icon-${size}.png`,
          sizes: `${size}x${size}`,
          type: "image/png",
          purpose: "any maskable",
        })),
      },
      workbox: {
        globPatterns: ["**/*.{js,css,html,png,webmanifest}"],
        navigateFallback: "index.html",
        cleanupOutdatedCaches: true,
        clientsClaim: true,
      },
    }),
  ],
  base: "./",
  build: {
    rolldownOptions: {
      input: ["index.html", "指数录入.html", "指数录入v0.91.html"],
      output: {
        manualChunks: (id) =>
          /node_modules.*(xlsx|fflate)/.test(id) ? "excel" : undefined,
      },
    },
  },
});
