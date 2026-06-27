export const metadata = {
  title: "St. Kizito's Primary School — Result Management System",
  description: "Result Management System for St. Kizito's Primary School",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body style={{ margin: 0, fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, sans-serif" }}>
        {children}
      </body>
    </html>
  );
}
