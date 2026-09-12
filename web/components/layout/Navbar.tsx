"use client";
import React, { useState } from "react";
import Link from "next/link";
import { usePathname } from "next/navigation";
import { motion, AnimatePresence } from "framer-motion";
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
  Menu,
  X,
  ChevronDown,
  ArrowRight,
} from "lucide-react";
import { useSeating } from "@/lib/context/SeatingContext";

export const Navbar = () => {
  const pathname = usePathname();
  const [mobileOpen, setMobileOpen] = useState(false);
  const [moreDropdownOpen, setMoreDropdownOpen] = useState(false);
  const [hoveredIdx, setHoveredIdx] = useState<number | null>(null);

  const {
    resetToDefaults,
    examMode,
    setExamMode,
    activeSession,
  } = useSeating();

  const mainNavLinks = [
    { name: "Dashboard", href: "/dashboard", icon: <LayoutDashboard className="h-4 w-4" /> },
    { name: "Seating Studio", href: "/studio", icon: <Grid3X3 className="h-4 w-4" /> },
    { name: "QR Scanner", href: "/scanner", icon: <QrCode className="h-4 w-4" /> },
    { name: "Proctors", href: "/invigilators", icon: <UserCheck className="h-4 w-4" /> },
    { name: "Room Builder", href: "/builder", icon: <Layers className="h-4 w-4" /> },
    { name: "Export", href: "/export", icon: <Printer className="h-4 w-4" /> },
  ];

  const extraLinks = [
    { name: "Student Kiosk", href: "/find-seat", icon: <Search className="h-4 w-4" /> },
    { name: "3D Hall Twin", href: "/3d-twin", icon: <Box className="h-4 w-4" /> },
  ];

  return (
    <header className="sticky top-0 z-50 w-full pt-3 px-4 sm:px-6 lg:px-8 pointer-events-none">
      <div className="max-w-7xl mx-auto pointer-events-auto">
        {/* Main Floating Glassmorphic Container */}
        <div className="h-16 px-4 sm:px-6 rounded-2xl bg-white/90 backdrop-blur-xl border border-[#E6E1D6] shadow-sm flex items-center justify-between gap-3">
          {/* Brand Logo */}
          <Link
            href="/"
            onClick={() => setMobileOpen(false)}
            className="flex items-center gap-2.5 group flex-shrink-0"
          >
            <div className="h-9 w-9 rounded-xl bg-gradient-to-tr from-emerald-600 to-teal-500 flex items-center justify-center shadow-md shadow-emerald-600/20 group-hover:scale-105 transition duration-200">
              <Sparkles className="h-4 w-4 text-white" />
            </div>
            <div className="flex flex-col">
              <div className="font-extrabold text-base tracking-tight flex items-center gap-1 leading-none text-slate-900">
                <span>Smart</span>
                <span className="text-gradient">Seating</span>
              </div>
              <span className="text-[10px] text-emerald-700 font-mono font-semibold tracking-wider uppercase mt-0.5">
                {examMode === "MST" ? "MST Test Mode" : "End-Sem Mode"}
              </span>
            </div>
          </Link>

          {/* Desktop Navigation Links */}
          <nav className="hidden lg:flex items-center gap-1 bg-[#F5F2EB] p-1 rounded-full border border-[#E4DFD4]">
            {mainNavLinks.map((item, idx) => {
              const isActive = pathname === item.href;
              return (
                <Link
                  key={item.href}
                  href={item.href}
                  onMouseEnter={() => setHoveredIdx(idx)}
                  onMouseLeave={() => setHoveredIdx(null)}
                  className={`relative text-xs font-semibold px-3 py-1.5 rounded-full transition duration-150 flex items-center gap-1.5 z-10 ${
                    isActive
                      ? "text-emerald-950 font-bold"
                      : "text-slate-600 hover:text-slate-900"
                  }`}
                >
                  {/* Active Animated Pill Indicator */}
                  {isActive && (
                    <motion.div
                      layoutId="navActivePill"
                      transition={{ type: "spring", stiffness: 450, damping: 32 }}
                      className="absolute inset-0 bg-white rounded-full shadow-xs border border-emerald-300/80 -z-10"
                    />
                  )}

                  {/* Hover Highlight */}
                  <AnimatePresence>
                    {hoveredIdx === idx && !isActive && (
                      <motion.div
                        layoutId="navHoverPill"
                        initial={{ opacity: 0, scale: 0.95 }}
                        animate={{ opacity: 1, scale: 1 }}
                        exit={{ opacity: 0, scale: 0.95 }}
                        className="absolute inset-0 bg-white/70 rounded-full -z-10"
                      />
                    )}
                  </AnimatePresence>

                  <span className={isActive ? "text-emerald-600" : "text-slate-500"}>
                    {item.icon}
                  </span>
                  <span>{item.name}</span>
                </Link>
              );
            })}

            {/* Extra Dropdown for Kiosk and 3D Twin */}
            <div className="relative">
              <button
                onClick={() => setMoreDropdownOpen(!moreDropdownOpen)}
                onBlur={() => setTimeout(() => setMoreDropdownOpen(false), 200)}
                className={`text-xs font-semibold px-3 py-1.5 rounded-full flex items-center gap-1 transition ${
                  pathname === "/find-seat" || pathname === "/3d-twin"
                    ? "bg-white text-emerald-950 font-bold shadow-xs border border-emerald-300/80"
                    : "text-slate-600 hover:text-slate-900"
                }`}
              >
                <span>More</span>
                <ChevronDown className="h-3 w-3 text-slate-400" />
              </button>

              <AnimatePresence>
                {moreDropdownOpen && (
                  <motion.div
                    initial={{ opacity: 0, y: 6 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: 6 }}
                    className="absolute right-0 top-full mt-2 w-44 bg-white border border-[#E6E1D6] rounded-2xl shadow-lg p-1.5 space-y-1 z-50"
                  >
                    {extraLinks.map((ex) => (
                      <Link
                        key={ex.href}
                        href={ex.href}
                        onClick={() => setMoreDropdownOpen(false)}
                        className={`flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-semibold transition ${
                          pathname === ex.href
                            ? "bg-emerald-50 text-emerald-900 font-bold"
                            : "text-slate-600 hover:bg-[#FAF8F3] hover:text-slate-900"
                        }`}
                      >
                        <span className="text-emerald-600">{ex.icon}</span>
                        <span>{ex.name}</span>
                      </Link>
                    ))}
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </nav>

          {/* Right Action Controls */}
          <div className="flex items-center gap-2.5 flex-shrink-0">
            {/* Mode Switcher Pill (MST vs End-Sem) */}
            <div className="hidden sm:flex items-center bg-[#F4EFE6] p-0.5 rounded-xl border border-[#E2DCCE]">
              <button
                onClick={() => setExamMode("MST")}
                className={`px-2.5 py-1 rounded-lg text-xs font-bold font-mono transition ${
                  examMode === "MST"
                    ? "bg-emerald-600 text-white shadow-xs"
                    : "text-slate-600 hover:text-slate-900"
                }`}
              >
                MST
              </button>
              <button
                onClick={() => setExamMode("END_SEM")}
                className={`px-2.5 py-1 rounded-lg text-xs font-bold font-mono transition ${
                  examMode === "END_SEM"
                    ? "bg-emerald-600 text-white shadow-xs"
                    : "text-slate-600 hover:text-slate-900"
                }`}
              >
                End-Sem
              </button>
            </div>

            {/* Reset Button */}
            <button
              onClick={resetToDefaults}
              title="Reset Demo Data"
              className="hidden md:flex items-center gap-1 p-2 rounded-xl text-slate-500 hover:text-emerald-800 hover:bg-[#FAF8F3] border border-transparent hover:border-[#E2DCCE] transition"
            >
              <RotateCcw className="h-4 w-4" />
            </button>

            {/* Studio Launch CTA */}
            <Link
              href="/studio"
              className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-md shadow-emerald-600/20"
            >
              <span>Launch Studio</span>
              <ArrowRight className="h-3.5 w-3.5" />
            </Link>

            {/* Mobile Hamburger Toggle */}
            <button
              onClick={() => setMobileOpen(!mobileOpen)}
              className="lg:hidden p-2 rounded-xl bg-[#FAF8F3] hover:bg-[#F3EFE6] border border-[#E0D9CB] text-slate-700 transition"
              aria-label="Toggle menu"
            >
              {mobileOpen ? <X className="h-5 w-5" /> : <Menu className="h-5 w-5" />}
            </button>
          </div>
        </div>

        {/* Mobile Animated Dropdown */}
        <AnimatePresence>
          {mobileOpen && (
            <motion.div
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="lg:hidden mt-2 bg-white/95 backdrop-blur-xl border border-[#E6E1D6] rounded-2xl p-4 shadow-xl space-y-3"
            >
              {/* Mobile Mode Switcher */}
              <div className="flex items-center justify-between p-2 rounded-xl bg-[#FAF8F3] border border-[#E8E2D4]">
                <span className="text-xs font-bold text-slate-700">Exam Mode:</span>
                <div className="flex items-center gap-1 bg-[#F3EFE6] p-0.5 rounded-lg border border-[#E0D9CB]">
                  <button
                    onClick={() => setExamMode("MST")}
                    className={`px-3 py-1 rounded text-xs font-bold font-mono transition ${
                      examMode === "MST" ? "bg-emerald-600 text-white" : "text-slate-600"
                    }`}
                  >
                    MST
                  </button>
                  <button
                    onClick={() => setExamMode("END_SEM")}
                    className={`px-3 py-1 rounded text-xs font-bold font-mono transition ${
                      examMode === "END_SEM" ? "bg-emerald-600 text-white" : "text-slate-600"
                    }`}
                  >
                    End-Sem
                  </button>
                </div>
              </div>

              {/* Mobile Links */}
              <div className="grid grid-cols-2 gap-2">
                {[...mainNavLinks, ...extraLinks].map((link) => {
                  const isActive = pathname === link.href;
                  return (
                    <Link
                      key={link.href}
                      href={link.href}
                      onClick={() => setMobileOpen(false)}
                      className={`flex items-center gap-2 p-2.5 rounded-xl text-xs font-bold transition ${
                        isActive
                          ? "bg-emerald-50 text-emerald-900 border border-emerald-300"
                          : "bg-[#FAF8F3] text-slate-700 hover:bg-[#F3EFE6] border border-[#E8E2D4]"
                      }`}
                    >
                      <span className={isActive ? "text-emerald-600" : "text-slate-500"}>
                        {link.icon}
                      </span>
                      <span>{link.name}</span>
                    </Link>
                  );
                })}
              </div>

              {/* Reset on Mobile */}
              <button
                onClick={() => {
                  resetToDefaults();
                  setMobileOpen(false);
                }}
                className="w-full py-2 rounded-xl bg-[#FAF8F3] border border-[#E0D9CB] text-xs font-bold text-slate-600 flex items-center justify-center gap-1.5"
              >
                <RotateCcw className="h-3.5 w-3.5" />
                <span>Reset Demo Data</span>
              </button>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </header>
  );
};
