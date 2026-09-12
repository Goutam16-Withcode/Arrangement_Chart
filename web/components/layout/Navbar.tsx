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
    <header className="sticky top-0 z-50 w-full glass-panel border-b border-slate-800/80">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
        {/* Brand Logo */}
        <Link href="/" className="flex items-center gap-3 group">
          <div className="h-10 w-10 rounded-xl bg-gradient-to-tr from-indigo-600 via-purple-600 to-pink-500 flex items-center justify-center shadow-lg shadow-indigo-500/25 group-hover:scale-105 transition duration-300">
            <Sparkles className="h-5 w-5 text-white" />
          </div>
          <div>
            <div className="font-bold text-lg tracking-tight flex items-center gap-1.5">
              <span className="text-white">Smart</span>
              <span className="text-gradient">Seating</span>
            </div>
            <div className="text-[10px] text-slate-400 font-mono tracking-wider uppercase -mt-0.5">
              Arrangement Studio
            </div>
          </div>
        </Link>

        {/* Navigation links */}
        <nav className="hidden md:flex items-center gap-1 bg-slate-900/60 p-1 rounded-xl border border-slate-800">
          {navLinks.map((link) => {
            const Icon = link.icon;
            const isActive = pathname === link.href;
            return (
              <Link
                key={link.href}
                href={link.href}
                className={`flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-xs font-medium transition duration-200 ${
                  isActive
                    ? "bg-indigo-600 text-white shadow-md shadow-indigo-600/30"
                    : "text-slate-300 hover:text-white hover:bg-slate-800/60"
                }`}
              >
                <Icon className="h-3.5 w-3.5" />
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
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs text-slate-400 hover:text-slate-200 bg-slate-800/40 hover:bg-slate-800 border border-slate-700/50 transition"
          >
            <RotateCcw className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">Reset Demo</span>
          </button>

          <Link
            href="/studio"
            className="relative inline-flex h-9 overflow-hidden rounded-xl p-[1px] focus:outline-none"
          >
            <span className="absolute inset-[-1000%] animate-[spin_3s_linear_infinite] bg-[conic-gradient(from_90deg_at_50%_50%,#E2CBFF_0%,#393BB2_50%,#E2CBFF_100%)]" />
            <span className="inline-flex h-full w-full cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-3.5 py-1 text-xs font-medium text-white backdrop-blur-3xl hover:bg-slate-900 transition">
              Launch Studio ⚡
            </span>
          </Link>
        </div>
      </div>
    </header>
  );
};
