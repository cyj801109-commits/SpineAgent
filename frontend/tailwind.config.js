/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      fontFamily: {
        // 루트 wrapper의 font-sans 클래스가 index.css의 body{font-family:'Pretendard'} 규칙보다
        // 우선 적용되어 실제로는 Tailwind 기본 sans 스택(Pretendard 미포함)이 렌더링되고 있었다.
        // font-sans가 Pretendard를 가리키도록 명시해 전체 사이트에 의도한 폰트가 적용되게 한다.
        sans: ['Pretendard', 'ui-sans-serif', 'system-ui', 'sans-serif'],
      },
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
