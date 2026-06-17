/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx,ts,tsx}"],
  theme: {
    extend: {
      colors: {
        appBg: "#f8f9fb",
        cardBg: "#ffffff",
        borderSoft: "#e2e8f0",
        textMain: "#1a202c",
        textMuted: "#718096",
        accentBlue: "#2563eb",
        accentGreen: "#10b981",
        accentAmber: "#f59e0b",
        accentOrange: "#f97316",
        accentPurple: "#8b5cf6",
      },
      boxShadow: {
        soft: "0 18px 48px rgba(26, 32, 44, 0.08)",
      },
      fontFamily: {
        display: ["Space Grotesk", "sans-serif"],
        body: ["Manrope", "sans-serif"],
      },
    },
  },
  plugins: [],
};
