import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import { AuthProvider } from "./components/auth-provider";
import { AuthenticatedAppFrame } from "./components/authenticated-app-frame";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "Event-Based Reminders",
  description: "Event-Based Reminders app",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html
      lang="en"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
    >
      <body className="min-h-full bg-gray-50 font-sans text-gray-900">
        <AuthProvider>
          <AuthenticatedAppFrame>{children}</AuthenticatedAppFrame>
        </AuthProvider>
      </body>
    </html>
  );
}
