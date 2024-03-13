import { resolve } from "path";
import { defineConfig } from "vite";
import dts from "vite-plugin-dts";
import react from "@vitejs/plugin-react";

// https://vitejs.dev/config/
export default defineConfig({
  build: {
    target: ["es2020"],
    lib: {
      entry: resolve(__dirname, "src/main.tsx"),
      name: "xlsx-merge-ts",
      fileName: (format) => `xlsx-merge-ts.${format}.js`,
    },
    rollupOptions: {
      external: ["react", "react-dom"],
      output: {
        globals: {
          react: "React",
          "react-dom": "ReactDOM",
        },
      },
    },
  },
  plugins: [
    react(),
    dts({
      insertTypesEntry: true,
    }),
  ],
  optimizeDeps: {
    exclude: ["three-gpu-pathtracer", "sqlite-wasm-esm"],
  },
});
