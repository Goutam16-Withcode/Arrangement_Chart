"use client";
import React from "react";
import Link from "next/link";
import { usePathname } from "next/navigation";
import {
  LayoutDashboard,
  Grid3X3,
  Layers,
  Search,
  Printer,
  Sparkles,
  RotateCcw,
  QrCode,
  UserCheck,
  Box,
} from "lucide-react";
import { useSeating } from "@/lib/context/SeatingContext";

export const Navbar = () => {
  const pathname = usePathname();
  const { resetToDefaults } = useSeating();

  const navLinks = [
    { name: "Dashboard", href: "/dashboard", icon: LayoutDashboard },
    { name: "Studio", href: "/studio", icon: Grid3X3 },
    { name: "QR Scanner", href: "/scanner", icon: QrCode },
    { name: "Proctor Roster", href: "/invigilators", icon: UserCheck },
    { name: "3D Twin", href: "/3d-twin", icon: Box },
    { name: "Builder", href: "/builder", icon: Layers },
    { name: "Student Kiosk", href: "/find-seat", icon: Search },
    { name: "Export", href: "/export", icon: Printer },
  ];

  return (
    <header className="sticky top-0 z-50 w-full bg-white/95 backdrop-blur-md border-b border-[#E8E2D4] shadow-xs">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
        {/* Brand Logo */}
        <Link href="/" className="flex items-center gap-3 group">
          <div className="h-10 w-10 rounded-xl bg-gradient-to-tr from-emerald-600 via-teal-600 to-emerald-400 flex items-center justify-center shadow-md shadow-emerald-500/20 group-hover:scale-105 transition duration-300">
            <Sparkles className="h-5 w-5 text-white" />
          </div>
          <div>
            <div className="font-bold text-lg tracking-tight flex items-center gap-1.5">
              <span className="text-slate-900">Smart</span>
              <span className="text-gradient">Seating</span>
            </div>
            <div className="text-[10px] text-emerald-700 font-mono tracking-wider uppercase -mt-0.5">
              Arrangement Studio
            </div>
          </div>
        </Link>

        {/* Navigation links */}
        <nav className="hidden lg:flex items-center gap-1 bg-[#F5F2EA] p-1 rounded-xl border border-[#E5DFD1]">
          {navLinks.map((link) => {
            const Icon = link.icon;
            const isActive = pathname === link.href;
            return (
              <Link
                key={link.href}
                href={link.href}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition duration-200 ${
                  isActive
                    ? "bg-white text-emerald-800 shadow-xs border border-emerald-100 font-bold"
                    : "text-slate-600 hover:text-emerald-800 hover:bg-white/60"
                }`}
              >
                <Icon className={`h-3.5 w-3.5 ${isActive ? "text-emerald-600" : "text-slate-500"}`} />
                <span>{link.name}</span>
              </Link>
            );
          })}
        </nav>

        {/* Right side actions */}
        <div className="flex items-center gap-2.5">
          <button
            onClick={resetToDefaults}
            title="Reset to Sample Demo Data"
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs text-slate-600 hover:text-emerald-800 bg-[#F4EFE6] hover:bg-[#EBE5DA] border border-[#E0D9CB] transition"
          >
            <RotateCcw className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">Reset</span>
          </button>

          <Link
            href="/scanner"
            className="px-3.5 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-xs"
          >
            <QrCode className="h-3.5 w-3.5" />
            <span>Scan QR</span>
          </Link>
        </div>
      </div>
    </header>
  );
};
