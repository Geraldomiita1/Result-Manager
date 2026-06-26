// ==============================================
// File: app/layout.js
// St. Kizito's Primary School MIS
// ==============================================

export const metadata = {
  title: "St. Kizito's Primary School — Result Management System",
  description: "Result Management System for St. Kizito's Primary School",
  applicationName: "St. Kizito's Primary School MIS",
  authors: [
    {
      name: "Gerald Omiita",
    },
  ],
  keywords: [
    "School MIS",
    "Result Management System",
    "Primary School",
    "Next.js",
    "Supabase",
    "Vercel",
    "Uganda",
    "St. Kizito's Primary School",
  ],
};

export const viewport = {
  width: "device-width",
  initialScale: 1,
  maximumScale: 1,
};

export default function RootLayout({ children }) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body
        style={{
          margin: 0,
          padding: 0,
          fontFamily:
            "system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
          backgroundColor: "#f4f6f8",
          color: "#222",
        }}
      >
        {children}
      </body>
    </html>
  );
}
