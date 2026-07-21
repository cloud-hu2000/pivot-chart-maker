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
        excelPivotTableOnline: `${projectRoot}/excel-pivot-table-online.html`,
        csvToBarChart: `${projectRoot}/csv-to-bar-chart.html`,
        csvToLineChart: `${projectRoot}/csv-to-line-chart.html`,
        csvToPieChart: `${projectRoot}/csv-to-pie-chart.html`,
        excelToBarChart: `${projectRoot}/excel-to-bar-chart.html`,
        excelToLineChart: `${projectRoot}/excel-to-line-chart.html`,
        excelToPieChart: `${projectRoot}/excel-to-pie-chart.html`,
      },
    },
  },
});
