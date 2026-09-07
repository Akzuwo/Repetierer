module.exports = {
  content: [
    "./public/**/*.html",
    "./public/**/*.js"
  ],
  theme: {
    extend: {
      colors: {
        ink: "#15171c",
        panel: "#20232a",
        mint: "#3b82f6",
        ember: "#93c5fd"
      },
      boxShadow: {
        glow: "0 0 0 1px rgba(59, 130, 246, 0.18), 0 24px 70px rgba(0, 0, 0, 0.44)",
        soft: "0 18px 42px rgba(0, 0, 0, 0.28)"
      }
    }
  },
  plugins: []
};
