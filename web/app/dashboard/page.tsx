"use client";
import React, { useState } from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import {
  Zap,
  ShieldCheck,
  Percent,
  Activity,
  Moon,
  ChevronDown,
  MoreVertical,
  ArrowUpRight,
  Sparkles,
  Layers,
  GraduationCap,
  Clock,
  BookOpen,
  CheckCircle2,
  Calendar,
} from "lucide-react";
import Link from "next/link";

export default function DashboardPage() {
  const {
    metrics,
    invigilators,
    roomSeatings,
    collegeProfile,
    activeSession,
    setIsConfigModalOpen,
  } = useSeating();

  const [filterPeriod, setFilterPeriod] = useState("Today");

  // Dot matrix data for Hall Utilization card
  const dotMatrixRows = [
    [true, true, true, true, false, true, true],
    [true, true, true, false, true, true, true],
    [true, false, true, true, true, true, false],
    [true, true, true, true, true, false, true],
    [false, true, true, true, true, true, true],
  ];

  return (
    <div className="space-y-6">
      {/* Title Bar (matching Health Overview in screenshot) */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
        <div>
          <h1 className="text-3xl sm:text-4xl font-black text-slate-900 tracking-tight">
            Examination Overview
          </h1>
          <p className="text-xs text-slate-500 mt-1 font-medium">
            Take control of your exam seating & hall allocations today!
          </p>
        </div>

        <div className="flex items-center gap-2 self-start sm:self-auto">
          <span className="text-xs text-slate-400 font-mono">
            {activeSession.date}
          </span>
          <button
            onClick={() => setIsConfigModalOpen(true)}
            className="px-3.5 py-1.5 rounded-full bg-white border border-slate-200/80 text-xs font-bold text-slate-800 shadow-2xs hover:bg-slate-50 transition"
          >
            Today ⌵
          </button>
        </div>
      </div>

      {/* Main Bento Grid Row */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
        {/* Card 1: Energy Used -> Capacity & Seating Distribution (Left 5 Cols) */}
        <div className="lg:col-span-5 bento-card p-6 flex flex-col justify-between space-y-6">
          {/* Card Top Title & Menu */}
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
              <Zap className="h-4 w-4 text-slate-900 fill-slate-900" />
              <span>Seating Capacity Allocation</span>
            </div>
            <button className="text-slate-400 hover:text-slate-800">
              <MoreVertical className="h-4 w-4" />
            </button>
          </div>

          {/* Large Metric */}
          <div>
            <div className="flex items-baseline gap-2">
              <span className="text-4xl font-black text-slate-900 font-mono tracking-tight">
                {metrics.seatedStudents || "4,3k"}
              </span>
              <span className="px-2 py-0.5 rounded-full bg-[#D4F754] text-black text-[11px] font-bold">
                +5%
              </span>
            </div>
            <span className="text-xs text-slate-400 font-medium mt-0.5 block">
              candidates seated today
            </span>
          </div>

          {/* Venn / Overlapping Circles Visualization (as in screenshot) */}
          <div className="relative h-44 w-full flex items-center justify-center my-2">
            {/* Lilac Circle (CSE) */}
            <div className="absolute left-6 md:left-12 h-28 w-28 rounded-full bg-[#B8B5FF] flex flex-col items-center justify-center text-slate-900 font-bold shadow-sm z-10">
              <span className="text-lg font-black leading-none">2,6k</span>
              <span className="text-[10px] text-slate-800 font-medium">CSE</span>
            </div>

            {/* Dark Charcoal Circle (ME) */}
            <div className="absolute right-6 md:right-12 h-28 w-28 rounded-full bg-[#1E1E22] flex flex-col items-center justify-center text-white font-bold shadow-md z-10">
              <span className="text-lg font-black leading-none">1,2k</span>
              <span className="text-[10px] text-slate-300 font-medium">ME</span>
            </div>

            {/* Neon Lime Small Overlap Circle (ECE) */}
            <div className="absolute bottom-1 h-20 w-20 rounded-full bg-[#D4F754] flex flex-col items-center justify-center text-black font-bold shadow-md z-20 border-2 border-white">
              <span className="text-sm font-black leading-none">500</span>
              <span className="text-[9px] text-black/80 font-bold">ECE</span>
            </div>
          </div>

          {/* Progress Percent Breakdown Bars (as in screenshot) */}
          <div className="space-y-3 pt-2">
            {/* CSE */}
            <div className="flex items-center justify-between text-xs">
              <div className="flex items-center gap-2">
                <span className="font-black text-slate-900">45%</span>
                <div className="w-32 sm:w-44 h-2 rounded-full bg-slate-100 overflow-hidden">
                  <div className="h-full bg-[#B8B5FF] rounded-full w-[45%]" />
                </div>
              </div>
              <span className="text-[11px] font-semibold text-slate-500 flex items-center gap-1">
                Computer Science <span className="h-1.5 w-1.5 rounded-full bg-[#B8B5FF]" />
              </span>
            </div>

            {/* ME */}
            <div className="flex items-center justify-between text-xs">
              <div className="flex items-center gap-2">
                <span className="font-black text-slate-900">30%</span>
                <div className="w-32 sm:w-44 h-2 rounded-full bg-slate-100 overflow-hidden">
                  <div className="h-full bg-[#1E1E22] rounded-full w-[30%]" />
                </div>
              </div>
              <span className="text-[11px] font-semibold text-slate-500 flex items-center gap-1">
                Mechanical Engg <span className="h-1.5 w-1.5 rounded-full bg-[#1E1E22]" />
              </span>
            </div>

            {/* ECE */}
            <div className="flex items-center justify-between text-xs">
              <div className="flex items-center gap-2">
                <span className="font-black text-slate-900">25%</span>
                <div className="w-32 sm:w-44 h-2 rounded-full bg-slate-100 overflow-hidden">
                  <div className="h-full bg-[#D4F754] rounded-full w-[25%]" />
                </div>
              </div>
              <span className="text-[11px] font-semibold text-slate-500 flex items-center gap-1">
                Electronics <span className="h-1.5 w-1.5 rounded-full bg-[#D4F754]" />
              </span>
            </div>
          </div>
        </div>

        {/* Right 7 Cols: 4 Sub-Cards Grid */}
        <div className="lg:col-span-7 grid grid-cols-1 sm:grid-cols-2 gap-5">
          {/* Card 2: Heart Rate -> Zero Conflict Index */}
          <div className="bento-card p-5 flex flex-col justify-between">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                <ShieldCheck className="h-4 w-4 text-slate-900" />
                <span>Zero Conflict Score</span>
              </div>
              <button className="text-slate-400 hover:text-slate-800">
                <MoreVertical className="h-4 w-4" />
              </button>
            </div>

            <div className="my-4">
              <div className="flex items-baseline gap-2">
                <span className="text-4xl font-black text-slate-900 font-mono">
                  {metrics.conflictFreeRate}%
                </span>
                <span className="text-xs text-slate-400 font-medium">
                  0 Conflicts
                </span>
              </div>
              <div className="text-[11px] text-slate-500 mt-1 flex items-center gap-1">
                <span className="h-2 w-2 rounded-full bg-[#D4F754]" />
                <span>Optimal Multi-Branch Separation</span>
              </div>
            </div>

            <div className="text-[10px] font-mono text-slate-400 pt-2 border-t border-slate-100 flex justify-between">
              <span>Avg 100% Target</span>
              <span>Constraint Engine Active</span>
            </div>
          </div>

          {/* Card 3: Wellness Index -> Hall Utilization + Dot Matrix Grid */}
          <div className="bento-card p-5 flex flex-col justify-between">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                <Percent className="h-4 w-4 text-slate-900" />
                <span>Hall Utilization</span>
              </div>
              <button className="text-slate-400 hover:text-slate-800">
                <MoreVertical className="h-4 w-4" />
              </button>
            </div>

            <div className="my-2 flex items-center justify-between">
              <div>
                <div className="flex items-baseline gap-1.5">
                  <span className="text-3xl font-black text-slate-900 font-mono">
                    {metrics.utilizationRate || 94}%
                  </span>
                  <span className="px-1.5 py-0.5 rounded-full bg-[#D4F754] text-black text-[10px] font-bold">
                    +10%
                  </span>
                </div>
                <span className="text-[11px] text-slate-400">Desk Efficiency</span>
              </div>

              {/* Dot Matrix Grid visualization */}
              <div className="grid grid-rows-5 gap-1">
                {dotMatrixRows.map((row, rIdx) => (
                  <div key={rIdx} className="flex gap-1">
                    {row.map((active, cIdx) => (
                      <span
                        key={cIdx}
                        className={`h-2 w-2 rounded-full ${
                          active
                            ? (rIdx + cIdx) % 3 === 0
                              ? "bg-[#D4F754]"
                              : "bg-[#B8B5FF]"
                            : "bg-slate-200"
                        }`}
                      />
                    ))}
                  </div>
                ))}
              </div>
            </div>

            <div className="text-[10px] font-mono text-slate-400 pt-2 border-t border-slate-100 flex justify-between">
              <span>Optimal Seat Density</span>
              <span>No Overcrowding</span>
            </div>
          </div>

          {/* Card 4: Activity -> Active Examination Halls */}
          <div className="bento-card p-5 flex flex-col justify-between">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                <Activity className="h-4 w-4 text-slate-900" />
                <span>Active Exam Halls</span>
              </div>
              <button className="text-slate-400 hover:text-slate-800">
                <MoreVertical className="h-4 w-4" />
              </button>
            </div>

            <div className="my-3">
              <div className="flex items-baseline gap-2">
                <span className="text-3xl font-black text-slate-900 font-mono">
                  {metrics.totalRooms || 3}
                </span>
                <span className="text-xs text-slate-500 font-bold">Halls Live</span>
              </div>
              <div className="text-[11px] text-slate-500 mt-1">
                75 Seats / Hall Average
              </div>
            </div>

            <div className="text-[10px] font-mono text-slate-400 pt-2 border-t border-slate-100 flex justify-between">
              <span>Rooms: 302, 304, 101</span>
              <Link href="/studio" className="text-black font-bold hover:underline">
                View Halls →
              </Link>
            </div>
          </div>

          {/* Card 5: Proctors Assigned */}
          <div className="bento-card p-5 flex flex-col justify-between">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                <GraduationCap className="h-4 w-4 text-slate-900" />
                <span>Invigilators on Duty</span>
              </div>
              <button className="text-slate-400 hover:text-slate-800">
                <MoreVertical className="h-4 w-4" />
              </button>
            </div>

            <div className="my-3 flex items-center gap-3">
              <div className="flex -space-x-2 overflow-hidden">
                {invigilators.slice(0, 3).map((inv) => (
                  <img
                    key={inv.id}
                    src={inv.image}
                    alt={inv.name}
                    className="inline-block h-8 w-8 rounded-full ring-2 ring-white object-cover"
                  />
                ))}
              </div>
              <div>
                <span className="text-sm font-black text-slate-900 block">
                  {invigilators.length} Assigned
                </span>
                <span className="text-[10px] text-slate-400">Anti-Bias Rotation</span>
              </div>
            </div>

            <div className="text-[10px] font-mono text-slate-400 pt-2 border-t border-slate-100 flex justify-between">
              <span>Chief: {collegeProfile.chiefSuperintendent}</span>
            </div>
          </div>

          {/* Card 6: Sleep Analysis -> Allocation Efficiency & Striped Timeline (Full width dark card) */}
          <div className="sm:col-span-2 bento-card-dark p-6 flex flex-col justify-between space-y-5">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-white">
                <Moon className="h-4 w-4 text-[#D4F754]" />
                <span>Hall Allocation Efficiency & Timelines</span>
              </div>

              <div className="px-3 py-1 rounded-full bg-white/10 text-white text-xs font-bold flex items-center gap-1 border border-white/10">
                <span>Session Timeline ⌵</span>
              </div>
            </div>

            {/* Metrics Row */}
            <div className="flex items-center gap-8">
              <div className="flex items-center gap-2.5">
                <div className="w-2.5 h-7 rounded-full bg-[#D4F754]" />
                <div>
                  <div className="text-2xl font-black text-white font-mono leading-none">
                    98%
                  </div>
                  <div className="text-[10px] text-slate-400 font-medium mt-0.5">
                    Seating Efficiency
                  </div>
                </div>
              </div>

              <div className="flex items-center gap-2.5">
                <div className="w-2.5 h-7 rounded-full bg-[#B8B5FF]" />
                <div>
                  <div className="text-2xl font-black text-white font-mono leading-none">
                    3h 00m
                  </div>
                  <div className="text-[10px] text-slate-400 font-medium mt-0.5">
                    Exam Duration
                  </div>
                </div>
              </div>
            </div>

            {/* Striped Bar Chart (matching Sleep Analysis bar chart in screenshot) */}
            <div className="pt-2">
              <div className="flex items-end justify-between gap-3 h-28 px-2">
                {/* Room 101 */}
                <div className="flex-1 flex flex-col items-center gap-2 h-full justify-end">
                  <div className="w-full bg-[#27272A] bg-striped-pattern h-[45%] rounded-xl" />
                  <span className="text-[10px] text-slate-400 font-mono">Rm 101</span>
                </div>

                {/* Room 202 */}
                <div className="flex-1 flex flex-col items-center gap-2 h-full justify-end">
                  <div className="w-full bg-[#27272A] bg-striped-pattern h-[60%] rounded-xl" />
                  <span className="text-[10px] text-slate-400 font-mono">Rm 202</span>
                </div>

                {/* Room 302 */}
                <div className="flex-1 flex flex-col items-center gap-2 h-full justify-end">
                  <div className="w-full bg-[#27272A] bg-striped-pattern h-[55%] rounded-xl" />
                  <span className="text-[10px] text-slate-400 font-mono">Rm 302</span>
                </div>

                {/* Room 304 (Active Highlighted Neon Lime & Lilac) */}
                <div className="flex-1 flex flex-col items-center gap-2 h-full justify-end">
                  <div className="w-full flex gap-1 h-full items-end">
                    <div className="w-1/2 bg-[#D4F754] h-[95%] rounded-xl shadow-lg shadow-[#D4F754]/20" />
                    <div className="w-1/2 bg-[#B8B5FF] h-[75%] rounded-xl" />
                  </div>
                  <span className="text-[10px] text-[#D4F754] font-bold font-mono flex items-center">
                    Hall 304 ↗
                  </span>
                </div>

                {/* Room 401 */}
                <div className="flex-1 flex flex-col items-center gap-2 h-full justify-end">
                  <div className="w-full bg-[#27272A] bg-striped-pattern h-[65%] rounded-xl" />
                  <span className="text-[10px] text-slate-400 font-mono">Rm 401</span>
                </div>

                {/* Room 402 */}
                <div className="flex-1 flex flex-col items-center gap-2 h-full justify-end">
                  <div className="w-full bg-[#27272A] bg-striped-pattern h-[40%] rounded-xl" />
                  <span className="text-[10px] text-slate-400 font-mono">Rm 402</span>
                </div>

                {/* Hall 501 */}
                <div className="flex-1 flex flex-col items-center gap-2 h-full justify-end">
                  <div className="w-full bg-[#27272A] bg-striped-pattern h-[50%] rounded-xl" />
                  <span className="text-[10px] text-slate-400 font-mono">Hall 501</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Quick Launch Action Ribbon */}
      <div className="bento-card p-5 flex flex-col sm:flex-row items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="h-10 w-10 rounded-2xl bg-[#D4F754] text-black flex items-center justify-center font-bold shadow-xs">
            <Sparkles className="h-5 w-5" />
          </div>
          <div>
            <h3 className="font-extrabold text-sm text-slate-900">
              Launch Visual 2D Seating Studio
            </h3>
            <p className="text-xs text-slate-500">
              Drag-and-drop seat swapping, anti-cheating check, and live QR code rosters.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          <Link
            href="/scanner"
            className="px-4 py-2 rounded-full bg-white hover:bg-slate-100 text-slate-800 border border-slate-200 text-xs font-bold transition shadow-2xs"
          >
            📷 QR Scanner
          </Link>
          <Link
            href="/studio"
            className="px-5 py-2 rounded-full bg-[#161618] hover:bg-black text-white text-xs font-bold transition shadow-sm flex items-center gap-1.5"
          >
            <span>Open Studio</span>
            <ArrowUpRight className="h-3.5 w-3.5" />
          </Link>
        </div>
      </div>
    </div>
  );
}
