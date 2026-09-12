import type { Metadata } from "next";
import "./globals.css";
import { AppShell } from "@/components/layout/AppShell";
import { SeatingProvider } from "@/lib/context/SeatingContext";

export const metadata: Metadata = {
  title: "flux.exam | Intelligent Exam Arrangement & Seating Studio",
  description:
    "Next-generation algorithmic exam seating chart generator, 2D visual classroom builder, and student kiosk pass system.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="min-h-screen bg-[#D2DEB1] text-slate-900 antialiased selection:bg-[#D4F754] selection:text-black">
        <SeatingProvider>
          <AppShell>{children}</AppShell>
        </SeatingProvider>
      </body>
    </html>
  );
}
