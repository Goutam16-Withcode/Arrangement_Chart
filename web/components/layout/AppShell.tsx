"use client";
import React, { useState, useEffect } from "react";
import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";
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
  FileSpreadsheet,
  CheckCircle2,
  DoorOpen,
  ArrowRight,
  ShieldCheck,
} from "lucide-react";
import { useSeating } from "@/lib/context/SeatingContext";
import { InstitutionModal } from "@/components/modals/InstitutionModal";
import confetti from "canvas-confetti";

export const AppShell: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const pathname = usePathname();
  const router = useRouter();
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [searchOpen, setSearchOpen] = useState(false);
  const [actionPaletteOpen, setActionPaletteOpen] = useState(false);

  const {
    collegeProfile,
    activeSession,
    examMode,
    setExamMode,
    setIsConfigModalOpen,
    resetToDefaults,
    students,
    roomSeatings,
    setSelectedRoomNumber,
    recalculateSeating,
  } = useSeating();

  const navItems = [
    {
      name: "Dashboard",
      href: "/dashboard",
      icon: <LayoutDashboard className="h-5 w-5" />,
      badge: "Live",
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
      badge: "Scan",
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
  ];

  // Instant Search Results
  const matchingStudents = searchQuery.trim()
    ? students
        .filter(
          (s) =>
            s.rollNo.toLowerCase().includes(searchQuery.toLowerCase()) ||
            s.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
            s.branch.toLowerCase().includes(searchQuery.toLowerCase())
        )
        .slice(0, 5)
    : [];

  const matchingRooms = searchQuery.trim()
    ? roomSeatings
        .filter(
          (r) =>
            r.roomConfig.roomNumber.toLowerCase().includes(searchQuery.toLowerCase()) ||
            r.roomConfig.building.toLowerCase().includes(searchQuery.toLowerCase())
        )
        .slice(0, 3)
    : [];

  const handleSelectStudent = (rollNo: string) => {
    // Find room where this student is allocated
    let foundRoom = "302";
    for (const rs of roomSeatings) {
      if (rs.attendanceList.some((a) => a.student.rollNo.toLowerCase() === rollNo.toLowerCase())) {
        foundRoom = rs.roomConfig.roomNumber;
        break;
      }
    }
    setSelectedRoomNumber(foundRoom);
    setSearchQuery("");
    setSearchOpen(false);
    router.push("/studio");
  };

  const handleSelectRoom = (roomNumber: string) => {
    setSelectedRoomNumber(roomNumber);
    setSearchQuery("");
    setSearchOpen(false);
    router.push("/studio");
  };

  const handleQuickOptimize = () => {
    recalculateSeating();
    confetti({
      particleCount: 50,
      spread: 80,
      origin: { y: 0.6 },
      colors: ["#D4F754", "#B8B5FF", "#161618"],
    });
    setActionPaletteOpen(false);
  };

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

        {/* Bottom Promo / Exam Optimization Widget Card */}
        <div className="mt-8 bg-[#D4F754] text-black p-4 rounded-3xl relative overflow-hidden shadow-lg space-y-3">
          <div className="flex items-center justify-between">
            <h4 className="font-black text-sm tracking-tight text-black">
              Anti-Cheat AI Engine
            </h4>
            <span className="text-base">⚡</span>
          </div>
          <p className="text-[11px] font-bold text-black/80 leading-snug">
            Automates zero adjacent branch conflict seating & attendance rosters.
          </p>
          <button
            onClick={handleQuickOptimize}
            className="w-full py-2.5 rounded-xl bg-[#161618] text-white text-xs font-black hover:bg-black transition shadow-sm text-center block"
          >
            Run Auto-Allocation
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
            className="p-2 rounded-xl bg-white/10 text-white text-xs font-bold"
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
                    ? "bg-white text-black font-extrabold"
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
              <div className="h-10 w-10 rounded-full bg-[#161618] text-[#D4F754] flex items-center justify-center font-black text-sm border-2 border-white shadow-xs overflow-hidden flex-shrink-0">
                <span>✦</span>
              </div>
              <div className="flex flex-col text-left">
                <div className="flex items-center gap-1.5">
                  <span className="font-black text-sm text-slate-900 group-hover:text-black transition">
                    {collegeProfile.collegeName}
                  </span>
                  <ChevronDown className="h-3.5 w-3.5 text-slate-400 group-hover:text-slate-900 transition" />
                </div>
                <span className="text-[11px] text-slate-500 font-mono font-bold">
                  {collegeProfile.collegeCode} • {activeSession.title}
                </span>
              </div>
            </div>

            {/* Right Controls: Search + Date Pill + Mode Switcher */}
            <div className="flex items-center gap-2.5 flex-wrap relative">
              {/* Search Bar with Instant Results Dropdown */}
              <div className="relative">
                <Search className="h-3.5 w-3.5 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="Search roll, hall, name..."
                  value={searchQuery}
                  onFocus={() => setSearchOpen(true)}
                  onChange={(e) => {
                    setSearchQuery(e.target.value);
                    setSearchOpen(true);
                  }}
                  className="pl-9 pr-4 py-2 rounded-full bg-white text-xs font-semibold text-slate-900 placeholder-slate-400 border border-slate-200/80 focus:outline-none focus:border-black shadow-2xs w-44 sm:w-60"
                />

                {/* Instant Search Dropdown Popover */}
                {searchOpen && searchQuery.trim().length > 0 && (
                  <div className="absolute right-0 top-11 w-72 sm:w-80 bg-white rounded-2xl shadow-xl border border-slate-200 p-3 z-50 space-y-3">
                    <div className="flex items-center justify-between text-[11px] font-bold text-slate-400 px-1 border-b border-slate-100 pb-1.5">
                      <span>Search Results</span>
                      <button
                        onClick={() => setSearchOpen(false)}
                        className="text-slate-400 hover:text-black"
                      >
                        ✕
                      </button>
                    </div>

                    {matchingStudents.length === 0 && matchingRooms.length === 0 && (
                      <div className="text-xs text-slate-500 py-3 text-center">
                        No candidates or halls matching &ldquo;{searchQuery}&rdquo;
                      </div>
                    )}

                    {/* Matching Candidates */}
                    {matchingStudents.length > 0 && (
                      <div className="space-y-1">
                        <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">
                          Candidates ({matchingStudents.length})
                        </span>
                        {matchingStudents.map((s) => (
                          <div
                            key={s.id}
                            onClick={() => handleSelectStudent(s.rollNo)}
                            className="p-2 rounded-xl hover:bg-slate-100 cursor-pointer flex items-center justify-between text-xs transition"
                          >
                            <div>
                              <div className="font-mono font-black text-slate-900">
                                {s.rollNo}
                              </div>
                              <div className="text-[10px] text-slate-500">{s.name} ({s.branch})</div>
                            </div>
                            <span className="text-[10px] font-bold bg-[#D4F754] text-black px-2 py-0.5 rounded-full">
                              View Pass →
                            </span>
                          </div>
                        ))}
                      </div>
                    )}

                    {/* Matching Halls */}
                    {matchingRooms.length > 0 && (
                      <div className="space-y-1 pt-1 border-t border-slate-100">
                        <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">
                          Halls ({matchingRooms.length})
                        </span>
                        {matchingRooms.map((r) => (
                          <div
                            key={r.roomConfig.roomNumber}
                            onClick={() => handleSelectRoom(r.roomConfig.roomNumber)}
                            className="p-2 rounded-xl hover:bg-slate-100 cursor-pointer flex items-center justify-between text-xs transition"
                          >
                            <div className="flex items-center gap-2">
                              <DoorOpen className="h-4 w-4 text-slate-600" />
                              <span className="font-bold text-slate-900">
                                Room {r.roomConfig.roomNumber} ({r.roomConfig.building})
                              </span>
                            </div>
                            <span className="text-[10px] font-mono text-slate-500">
                              {r.assignedCount}/{r.totalCapacity}
                            </span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}
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
          <div
            onClick={() => setActionPaletteOpen(true)}
            className="inline-flex items-center gap-2.5 px-4 py-2.5 rounded-full bg-[#161618] text-white text-xs font-black shadow-2xl border border-white/10 hover:scale-105 transition cursor-pointer group"
          >
            <Volume2 className="h-3.5 w-3.5 text-[#D4F754] group-hover:rotate-12 transition-transform" />
            <span>Action key mode • AI Seating Active</span>
            <span className="text-[10px] bg-white/20 px-2 py-0.5 rounded-full font-mono">
              Quick Menu ↗
            </span>
          </div>
        </div>
      </div>

      {/* Action Key Command Palette Modal */}
      <AnimatePresence>
        {actionPaletteOpen && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-[32px] max-w-lg w-full p-6 shadow-2xl border border-slate-200 space-y-5"
            >
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <div className="p-2 rounded-xl bg-[#161618] text-[#D4F754]">
                    <Zap className="h-5 w-5" />
                  </div>
                  <div>
                    <h3 className="font-black text-slate-900 text-base">
                      Exam Action Palette
                    </h3>
                    <p className="text-xs text-slate-500 font-medium">
                      Instant operations for Controller of Examinations
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setActionPaletteOpen(false)}
                  className="p-2 rounded-full hover:bg-slate-100 text-slate-500 font-bold"
                >
                  ✕
                </button>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <button
                  onClick={handleQuickOptimize}
                  className="p-3.5 rounded-2xl bg-slate-50 hover:bg-[#D4F754]/20 border border-slate-200 hover:border-black text-left transition flex items-center justify-between group"
                >
                  <div className="space-y-0.5">
                    <div className="text-xs font-black text-slate-900">
                      ⚡ Run Anti-Cheat Optimizer
                    </div>
                    <div className="text-[11px] text-slate-500">
                      Rebalance multi-branch seats
                    </div>
                  </div>
                  <ArrowRight className="h-4 w-4 text-slate-400 group-hover:text-black group-hover:translate-x-0.5 transition" />
                </button>

                <button
                  onClick={() => {
                    setActionPaletteOpen(false);
                    router.push("/export");
                  }}
                  className="p-3.5 rounded-2xl bg-slate-50 hover:bg-[#B8B5FF]/20 border border-slate-200 hover:border-black text-left transition flex items-center justify-between group"
                >
                  <div className="space-y-0.5">
                    <div className="text-xs font-black text-slate-900">
                      🖨️ Print Attendance Books
                    </div>
                    <div className="text-[11px] text-slate-500">
                      Section-wise registers & notices
                    </div>
                  </div>
                  <ArrowRight className="h-4 w-4 text-slate-400 group-hover:text-black group-hover:translate-x-0.5 transition" />
                </button>

                <button
                  onClick={() => {
                    setActionPaletteOpen(false);
                    router.push("/scanner");
                  }}
                  className="p-3.5 rounded-2xl bg-slate-50 hover:bg-slate-100 border border-slate-200 hover:border-black text-left transition flex items-center justify-between group"
                >
                  <div className="space-y-0.5">
                    <div className="text-xs font-black text-slate-900">
                      📷 Launch QR Scanner
                    </div>
                    <div className="text-[11px] text-slate-500">
                      Hall entry biometric check
                    </div>
                  </div>
                  <ArrowRight className="h-4 w-4 text-slate-400 group-hover:text-black group-hover:translate-x-0.5 transition" />
                </button>

                <button
                  onClick={() => {
                    setActionPaletteOpen(false);
                    router.push("/builder");
                  }}
                  className="p-3.5 rounded-2xl bg-slate-50 hover:bg-slate-100 border border-slate-200 hover:border-black text-left transition flex items-center justify-between group"
                >
                  <div className="space-y-0.5">
                    <div className="text-xs font-black text-slate-900">
                      ➕ Add Examination Hall
                    </div>
                    <div className="text-[11px] text-slate-500">
                      2D Blueprint room creator
                    </div>
                  </div>
                  <ArrowRight className="h-4 w-4 text-slate-400 group-hover:text-black group-hover:translate-x-0.5 transition" />
                </button>
              </div>

              <div className="pt-2 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500 font-medium">
                <span>Active: {activeSession.title} ({activeSession.examMode})</span>
                <button
                  onClick={() => {
                    setActionPaletteOpen(false);
                    setIsConfigModalOpen(true);
                  }}
                  className="text-black font-bold hover:underline"
                >
                  Edit Session Settings →
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};
