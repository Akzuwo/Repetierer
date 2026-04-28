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
        mint: "#18a890",
        ember: "#f0c44c"
      },
      boxShadow: {
        glow: "0 0 0 1px rgba(24, 168, 144, 0.18), 0 24px 70px rgba(0, 0, 0, 0.34)",
        soft: "0 18px 42px rgba(0, 0, 0, 0.28)"
      }
    }
  },
  plugins: []
};
