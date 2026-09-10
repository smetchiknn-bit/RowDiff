/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      fontFamily: {
        sans: ['IBM Plex Sans', 'system-ui', 'sans-serif'],
        mono: ['IBM Plex Mono', 'monospace'],
        display: ['Unbounded', 'system-ui'],
      },
      colors: {
        green: {
          dark: '#123f28',
          primary: '#1e7145',
          light: '#e0efe5',
          fog: '#f0f7f2',
        },
        paper: '#eef2ec',
      },
      boxShadow: {
        'hard': '5px 6px 0 rgba(21,38,32,0.07)',
        'hard-sm': '3px 3px 0 rgba(0,0,0,0.25)',
      },
      animation: {
        'rise-in': 'riseIn 0.45s ease-out forwards',
        'cell-flash': 'cellFlash 0.6s ease-in-out',
        'dash-move': 'dashMove 2s linear infinite',
        'float-soft': 'floatSoft 10s ease-in-out infinite',
      },
      keyframes: {
        riseIn: {
          '0%': { opacity: '0', transform: 'translateY(14px)' },
          '100%': { opacity: '1', transform: 'translateY(0)' },
        },
        cellFlash: {
          '0%, 100%': { backgroundColor: 'transparent' },
          '50%': { backgroundColor: '#e0efe5' },
        },
        dashMove: {
          '0%': { strokeDashoffset: '0' },
          '100%': { strokeDashoffset: '-10' },
        },
        floatSoft: {
          '0%, 100%': { transform: 'translateY(0)' },
          '50%': { transform: 'translateY(-8px)' },
        },
      },
    },
  },
  plugins: [],
}
