import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import { fileURLToPath } from "node:url";

const projectRoot = fileURLToPath(new URL(".", import.meta.url));

export default defineConfig({
  plugins: [react()],
  build: {
    rollupOptions: {
      input: {
        main: `${projectRoot}/index.html`,
        pivotTableGenerator: `${projectRoot}/pivot-table-generator.html`,
      },
    },
  },
});
