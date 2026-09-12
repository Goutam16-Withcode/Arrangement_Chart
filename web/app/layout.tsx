import type { Metadata } from "next";
import "./globals.css";
import { Navbar } from "@/components/layout/Navbar";
import { SeatingProvider } from "@/lib/context/SeatingContext";

export const metadata: Metadata = {
  title: "Smart Seating Studio | Intelligent Exam Arrangement & Attendance",
  description:
    "Next-generation algorithmic exam seating chart generator, 2D visual classroom builder, and student kiosk pass system.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="dark">
      <body className="min-h-screen bg-[#090D16] text-slate-100 antialiased selection:bg-indigo-500/30 selection:text-indigo-200">
        <SeatingProvider>
          <div className="flex flex-col min-h-screen">
            <Navbar />
            <main className="flex-1">{children}</main>
            <footer className="glass-panel border-t border-slate-800/80 py-6 text-center text-xs text-slate-500 font-mono">
              Smart Exam Seating Arrangement Studio • Built with Next.js & Aceternity UI
            </footer>
          </div>
        </SeatingProvider>
      </body>
    </html>
  );
}
