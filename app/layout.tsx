import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "HROne Counter Builder",
  description:
    "Build and verify balanced SAP posting counters from monthly HROne journal extracts.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
