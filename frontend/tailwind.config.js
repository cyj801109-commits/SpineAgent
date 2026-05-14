/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        primary: '#333333',
        accent: '#00BCD4',
        sub: '#888888',
        borderline: '#E5E7EB',
        paper: '#FFFFFF',
        pagebg: '#F4F5F7'
      }
    },
  },
  plugins: [],
}
