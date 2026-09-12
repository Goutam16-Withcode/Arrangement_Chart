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
} from "lucide-react";
import { useSeating } from "@/lib/context/SeatingContext";

export const Navbar = () => {
  const pathname = usePathname();
  const { resetToDefaults } = useSeating();

  const navLinks = [
    { name: "Dashboard", href: "/dashboard", icon: LayoutDashboard },
    { name: "Seating Studio", href: "/studio", icon: Grid3X3 },
    { name: "Room Builder", href: "/builder", icon: Layers },
    { name: "Student Kiosk", href: "/find-seat", icon: Search },
    { name: "Export & Print", href: "/export", icon: Printer },
  ];

  return (
    <header className="sticky top-0 z-50 w-full bg-white/90 backdrop-blur-md border-b border-[#E8E2D4] shadow-sm">
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
        <nav className="hidden md:flex items-center gap-1 bg-[#F5F2EA] p-1 rounded-xl border border-[#E5DFD1]">
          {navLinks.map((link) => {
            const Icon = link.icon;
            const isActive = pathname === link.href;
            return (
              <Link
                key={link.href}
                href={link.href}
                className={`flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-xs font-semibold transition duration-200 ${
                  isActive
                    ? "bg-white text-emerald-800 shadow-sm border border-emerald-100"
                    : "text-slate-600 hover:text-emerald-800 hover:bg-white/60"
                }`}
              >
                <Icon className={`h-3.5 w-3.5 ${isActive ? "text-emerald-600" : "text-slate-500"}`} />
                {link.name}
              </Link>
            );
          })}
        </nav>

        {/* Right side actions */}
        <div className="flex items-center gap-3">
          <button
            onClick={resetToDefaults}
            title="Reset to Sample Demo Data"
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs text-slate-600 hover:text-emerald-800 bg-[#F4EFE6] hover:bg-[#EBE5DA] border border-[#E0D9CB] transition"
          >
            <RotateCcw className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">Reset Demo</span>
          </button>

          <Link
            href="/studio"
            className="relative inline-flex h-9 overflow-hidden rounded-xl p-[1px] focus:outline-none"
          >
            <span className="absolute inset-[-1000%] animate-[spin_3s_linear_infinite] bg-[conic-gradient(from_90deg_at_50%_50%,#A7F3D0_0%,#059669_50%,#A7F3D0_100%)]" />
            <span className="inline-flex h-full w-full cursor-pointer items-center justify-center rounded-xl bg-emerald-600 hover:bg-emerald-700 px-3.5 py-1 text-xs font-bold text-white shadow-md shadow-emerald-600/20 transition">
              Launch Studio ⚡
            </span>
          </Link>
        </div>
      </div>
    </header>
  );
};
