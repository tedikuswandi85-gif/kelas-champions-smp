import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  // base: "./" membuat semua path file CSS/JS bersifat relatif, sehingga aplikasi langsung
  // berjalan benar di GitHub Pages (URL berbentuk https://username.github.io/nama-repo/)
  // TANPA perlu mengedit file ini lagi untuk mencocokkan nama repo — cukup build & deploy.
  base: "./",
});
