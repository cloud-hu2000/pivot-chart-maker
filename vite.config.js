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
        tools: `${projectRoot}/tools.html`,
        chartMaker: `${projectRoot}/chart-maker.html`,
        csvViewer: `${projectRoot}/csv-viewer.html`,
        jsonToCsv: `${projectRoot}/json-to-csv.html`,
        csvToJson: `${projectRoot}/csv-to-json.html`,
        csvToExcel: `${projectRoot}/csv-to-excel.html`,
      },
    },
  },
});
