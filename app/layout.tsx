import type { Metadata } from "next";
import { headers } from "next/headers";
import "./globals.css";

const title = "HROne Counter Builder";
const description =
  "Build and verify balanced SAP posting counters from monthly HROne journal extracts.";

export async function generateMetadata(): Promise<Metadata> {
  const requestHeaders = await headers();
  const host =
    requestHeaders.get("x-forwarded-host") ?? requestHeaders.get("host");
  const protocol =
    requestHeaders.get("x-forwarded-proto") ??
    (host?.startsWith("localhost") ? "http" : "https");
  const origin = host
    ? `${protocol}://${host}`
    : "https://hrone-counter-builder-jain.deepak-basera1985.chatgpt.site";
  const imageUrl = `${origin}/og.png`;

  return {
    title,
    description,
    openGraph: {
      title,
      description,
      type: "website",
      url: origin,
      images: [
        {
          url: imageUrl,
          width: 1672,
          height: 941,
          alt: "HROne Counter Builder — balanced monthly journals",
        },
      ],
    },
    twitter: {
      card: "summary_large_image",
      title,
      description,
      images: [imageUrl],
    },
  };
}

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
