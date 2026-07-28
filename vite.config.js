import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Configuración de compilación. base: "./" hace que las rutas sean relativas,
// para que funcione igual en Netlify o en cualquier subcarpeta.
// minify: "terser" evita depender de esbuild para minificar, que en algunas
// instalaciones de Windows no se activa y hace fallar la compilación.
export default defineConfig({
  plugins: [react()],
  base: "./",
  build: {
    outDir: "dist",
    minify: "terser",
    chunkSizeWarningLimit: 1200,
  },
});
