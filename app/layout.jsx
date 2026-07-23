import "./globals.css";

export const metadata = {
  title: "Fundamental Screener",
  description: "Analisa Saham Indonesia",
};

export default function RootLayout({ children }) {
  return (
    <html lang="id">
      <body>{children}</body>
    </html>
  );
}