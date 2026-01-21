import netlify from '@netlify/vite-plugin'
import tailwindcss from '@tailwindcss/vite'
// https://vite.dev/config/
export default {
  plugins: [
    netlify(),
    tailwindcss(),
  ],
}
