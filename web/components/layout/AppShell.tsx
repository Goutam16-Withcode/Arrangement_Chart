"use client";
import React, { useState } from "react";
import Link from "next/link";
import { usePathname } from "next/navigation";
import { motion, AnimatePresence } from "framer-motion";
import {
  LayoutDashboard,
  Home,
  Grid3X3,
  QrCode,
  UserCheck,
  Layers,
  Printer,
  Search,
  Box,
  ChevronDown,
  Sparkles,
  Sliders,
  RotateCcw,
  Zap,
  Menu,
  X,
  Volume2,
  Calendar,
} from "lucide-react";
import { useSeating } from "@/lib/context/SeatingContext";
import { InstitutionModal } from "@/components/modals/InstitutionModal";

export const AppShell: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const pathname = usePathname();
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");

  const {
    collegeProfile,
    activeSession,
    examMode,
    setExamMode,
    setIsConfigModalOpen,
    resetToDefaults,
  } = useSeating();

  const navItems = [
    {
      name: "Dashboard",
      href: "/dashboard",
      icon: <LayoutDashboard className="h-5 w-5" />,
      badge: "3",
    },
    {
      name: "Home",
      href: "/",
      icon: <Home className="h-5 w-5" />,
    },
    {
      name: "Seating Studio",
      href: "/studio",
      icon: <Grid3X3 className="h-5 w-5" />,
    },
    {
      name: "QR Scanner",
      href: "/scanner",
      icon: <QrCode className="h-5 w-5" />,
      badge: "1",
    },
    {
      name: "Invigilators",
      href: "/invigilators",
      icon: <UserCheck className="h-5 w-5" />,
    },
    {
      name: "Room Builder",
      href: "/builder",
      icon: <Layers className="h-5 w-5" />,
    },
    {
      name: "Export Station",
      href: "/export",
      icon: <Printer className="h-5 w-5" />,
    },
    {
      name: "Student Kiosk",
      href: "/find-seat",
      icon: <Search className="h-5 w-5" />,
    },
    {
      name: "3D Hall Twin",
      href: "/3d-twin",
      icon: <Box className="h-5 w-5" />,
    },
  ];

  return (
    <div className="min-h-[92vh] max-w-[1580px] mx-auto app-outer-shell flex flex-col lg:flex-row relative">
      <InstitutionModal />

      {/* Desktop & Tablet Sidebar */}
      <aside className="hidden lg:flex w-64 bg-[#161618] p-5 flex-col justify-between flex-shrink-0 text-white rounded-l-[36px]">
        {/* Brand Logo */}
        <div className="space-y-6">
          <Link
            href="/dashboard"
            className="flex items-center gap-3 px-2 py-1 group"
          >
            <div className="h-9 w-9 rounded-2xl bg-[#D4F754] text-black flex items-center justify-center font-black text-xl shadow-md shadow-[#D4F754]/20 group-hover:scale-105 transition">
              ✦
            </div>
            <div className="flex flex-col">
              <span className="font-extrabold text-lg tracking-tight text-white flex items-center gap-1">
                <span>flux</span>
                <span className="text-[#D4F754] text-xs font-mono font-bold">.exam</span>
              </span>
            </div>
          </Link>

          {/* Navigation Links List */}
          <nav className="space-y-1.5">
            {navItems.map((item) => {
              const isActive = pathname === item.href;
              return (
                <Link
                  key={item.href}
                  href={item.href}
                  className={`flex items-center justify-between px-4 py-3 rounded-full text-xs font-bold transition-all duration-150 ${
                    isActive
                      ? "bg-white text-black shadow-md font-extrabold"
                      : "text-[#8E8E93] hover:text-white hover:bg-white/5"
                  }`}
                >
                  <div className="flex items-center gap-3">
                    <span className={isActive ? "text-black" : "text-[#8E8E93]"}>
                      {item.icon}
                    </span>
                    <span>{item.name}</span>
                  </div>

                  {item.badge && (
                    <span
                      className={`px-2 py-0.5 rounded-full text-[10px] font-mono font-bold ${
                        isActive
                          ? "bg-[#D4F754] text-black"
                          : "bg-[#D4F754] text-black"
                      }`}
                    >
                      {item.badge}
                    </span>
                  )}
                </Link>
              );
            })}
          </nav>
        </div>

        {/* Bottom Promo / Upgrade Widget Card (like the Rocket Card in the screenshot) */}
        <div className="mt-8 bg-[#D4F754] text-black p-4 rounded-3xl relative overflow-hidden shadow-lg space-y-3">
          <div className="flex items-center justify-between">
            <h4 className="font-black text-sm tracking-tight text-black">
              Universal Exam AI
            </h4>
            <span className="text-base">🚀</span>
          </div>
          <p className="text-[11px] font-semibold text-black/80 leading-snug">
            Configure infinite halls, anti-cheating separation & live QR attendance.
          </p>
          <button
            onClick={() => setIsConfigModalOpen(true)}
            className="w-full py-2.5 rounded-xl bg-[#161618] text-white text-xs font-bold hover:bg-black transition shadow-sm text-center block"
          >
            Configure Profile
          </button>
        </div>
      </aside>

      {/* Mobile Top Header (Small Screens) */}
      <div className="lg:hidden bg-[#161618] p-4 flex items-center justify-between text-white border-b border-white/10 rounded-t-[36px]">
        <Link href="/dashboard" className="flex items-center gap-2">
          <div className="h-8 w-8 rounded-xl bg-[#D4F754] text-black flex items-center justify-center font-black text-lg">
            ✦
          </div>
          <span className="font-extrabold text-base tracking-tight">flux.exam</span>
        </Link>

        <div className="flex items-center gap-2">
          <button
            onClick={() => setIsConfigModalOpen(true)}
            className="p-2 rounded-xl bg-white/10 text-white text-xs"
          >
            <Sliders className="h-4 w-4" />
          </button>
          <button
            onClick={() => setMobileSidebarOpen(!mobileSidebarOpen)}
            className="p-2 rounded-xl bg-white/10 text-white"
          >
            {mobileSidebarOpen ? <X className="h-5 w-5" /> : <Menu className="h-5 w-5" />}
          </button>
        </div>
      </div>

      {/* Mobile Sidebar Dropdown Drawer */}
      <AnimatePresence>
        {mobileSidebarOpen && (
          <motion.div
            initial={{ opacity: 0, height: 0 }}
            animate={{ opacity: 1, height: "auto" }}
            exit={{ opacity: 0, height: 0 }}
            className="lg:hidden bg-[#161618] border-b border-white/10 p-4 space-y-2 text-white overflow-hidden"
          >
            {navItems.map((item) => (
              <Link
                key={item.href}
                href={item.href}
                onClick={() => setMobileSidebarOpen(false)}
                className={`flex items-center justify-between p-3 rounded-2xl text-xs font-bold ${
                  pathname === item.href
                    ? "bg-white text-black"
                    : "text-slate-300 hover:bg-white/10"
                }`}
              >
                <div className="flex items-center gap-3">
                  {item.icon}
                  <span>{item.name}</span>
                </div>
                {item.badge && (
                  <span className="px-2 py-0.5 rounded-full bg-[#D4F754] text-black text-[10px] font-bold">
                    {item.badge}
                  </span>
                )}
              </Link>
            ))}
          </motion.div>
        )}
      </AnimatePresence>

      {/* Main Content Workspace (Light Gray Inner Card) */}
      <div className="flex-1 main-content-panel m-2.5 sm:m-3 p-4 sm:p-7 flex flex-col justify-between overflow-x-hidden min-h-[85vh]">
        <div className="space-y-6">
          {/* Top Floating App Bar */}
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-2 border-b border-slate-200/60">
            {/* User & Institution Profile Pill */}
            <div
              onClick={() => setIsConfigModalOpen(true)}
              className="flex items-center gap-3 cursor-pointer group p-1.5 rounded-2xl hover:bg-white transition"
            >
              <div className="h-10 w-10 rounded-full bg-slate-900 text-white flex items-center justify-center font-bold text-sm border-2 border-white shadow-xs overflow-hidden flex-shrink-0">
                <span className="text-[#D4F754] text-base">✦</span>
              </div>
              <div className="flex flex-col text-left">
                <div className="flex items-center gap-1.5">
                  <span className="font-extrabold text-sm text-slate-900 group-hover:text-black transition">
                    {collegeProfile.collegeName}
                  </span>
                  <ChevronDown className="h-3.5 w-3.5 text-slate-400 group-hover:text-slate-900 transition" />
                </div>
                <span className="text-[11px] text-slate-500 font-mono">
                  {collegeProfile.collegeCode} • {activeSession.title}
                </span>
              </div>
            </div>

            {/* Right Controls: Search + Date Pill + Mode Switcher */}
            <div className="flex items-center gap-2.5 flex-wrap">
              {/* Search Bar */}
              <div className="relative">
                <Search className="h-3.5 w-3.5 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="Search roll, hall, subject..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="pl-9 pr-4 py-2 rounded-full bg-white text-xs font-semibold text-slate-900 placeholder-slate-400 border border-slate-200/80 focus:outline-none focus:border-slate-400 shadow-2xs w-44 sm:w-56"
                />
              </div>

              {/* Mode Switcher Pill */}
              <div className="flex items-center bg-white p-1 rounded-full border border-slate-200/80 shadow-2xs">
                <button
                  onClick={() => setExamMode("MST")}
                  className={`px-3 py-1 rounded-full text-xs font-bold transition ${
                    examMode === "MST"
                      ? "bg-[#D4F754] text-black font-extrabold"
                      : "text-slate-500 hover:text-slate-900"
                  }`}
                >
                  MST
                </button>
                <button
                  onClick={() => setExamMode("END_SEM")}
                  className={`px-3 py-1 rounded-full text-xs font-bold transition ${
                    examMode === "END_SEM"
                      ? "bg-[#D4F754] text-black font-extrabold"
                      : "text-slate-500 hover:text-slate-900"
                  }`}
                >
                  End-Sem
                </button>
              </div>

              {/* Date / Today Tag */}
              <div className="hidden sm:flex items-center gap-1.5 px-3.5 py-2 rounded-full bg-white border border-slate-200/80 text-xs font-bold text-slate-700 shadow-2xs">
                <Calendar className="h-3.5 w-3.5 text-slate-500" />
                <span>{activeSession.date}</span>
              </div>

              {/* Reset Demo Data */}
              <button
                onClick={resetToDefaults}
                title="Reset Demo Data"
                className="p-2 rounded-full bg-white hover:bg-slate-100 text-slate-500 border border-slate-200/80 transition"
              >
                <RotateCcw className="h-3.5 w-3.5" />
              </button>
            </div>
          </div>

          {/* Render Active Route Body */}
          <main className="min-h-[70vh]">{children}</main>
        </div>

        {/* Bottom Floating Action Key Mode Capsule */}
        <div className="mt-8 flex justify-center">
          <div className="inline-flex items-center gap-2.5 px-4 py-2 rounded-2xl bg-[#18181B] text-white text-xs font-bold shadow-2xl border border-white/10 hover:scale-105 transition cursor-pointer">
            <Volume2 className="h-3.5 w-3.5 text-[#D4F754]" />
            <span>Action key mode • AI Seating Active</span>
          </div>
        </div>
      </div>
    </div>
  );
};
