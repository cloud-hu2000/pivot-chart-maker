import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import { fileURLToPath } from "node:url";

const projectRoot = fileURLToPath(new URL(".", import.meta.url));
const landingRewrites = {
  "/online-pivot-table": "/online-pivot-table.html",
  "/pivot-chart-without-excel": "/pivot-chart-without-excel.html",
  "/xlsx-to-pivot-chart": "/xlsx-to-pivot-chart.html",
};

function rewriteLandingPaths(req) {
  const [path, query] = String(req.url || "").split("?");
  const target = landingRewrites[path];
  if (target) req.url = query ? `${target}?${query}` : target;
}

export default defineConfig({
  plugins: [
    react(),
    {
      name: "landing-rewrites",
      configureServer(server) {
        server.middlewares.use((req, _res, next) => {
          rewriteLandingPaths(req);
          next();
        });
      },
      configurePreviewServer(server) {
        server.middlewares.use((req, _res, next) => {
          rewriteLandingPaths(req);
          next();
        });
      },
    },
  ],
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
        onlinePivotTable: `${projectRoot}/online-pivot-table.html`,
        pivotChartWithoutExcel: `${projectRoot}/pivot-chart-without-excel.html`,
        xlsxToPivotChart: `${projectRoot}/xlsx-to-pivot-chart.html`,
      },
    },
  },
});
