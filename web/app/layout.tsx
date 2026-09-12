import type { Metadata } from "next";
import "./globals.css";
import { Navbar } from "@/components/layout/Navbar";
import { InstitutionModal } from "@/components/modals/InstitutionModal";
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
    <html lang="en">
      <body className="min-h-screen bg-[#FBF9F4] text-slate-900 antialiased selection:bg-emerald-100 selection:text-emerald-900">
        <SeatingProvider>
          <div className="flex flex-col min-h-screen bg-[#FBF9F4]">
            <Navbar />
            <InstitutionModal />
            <main className="flex-1">{children}</main>
            <footer className="glass-panel border-t border-[#E8E2D4] py-6 text-center text-xs text-slate-500 font-mono">
              Smart Exam Seating Arrangement Studio • Universal Edition
            </footer>
          </div>
        </SeatingProvider>
      </body>
    </html>
  );
}
