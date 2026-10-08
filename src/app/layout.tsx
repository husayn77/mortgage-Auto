import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "MGT Mapper",
  description: "Map member details into the MGT Excel formula template.",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
