/** @type {import('tailwindcss').Config} */
export default {
  darkMode: "class", // 👈 important for manual dark mode toggle
  content: [
    "./index.html",
    "./src/**/*.{ts,tsx,js,jsx}", // all js/ts/tsx/jsx files
  ],
  theme: {
    extend: {},
  },
  plugins: [],
};
