import "./globals.css";

export const metadata = {
  title: "Mail Summary View",
};

export default function RootLayout({ children }) {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  );
}
