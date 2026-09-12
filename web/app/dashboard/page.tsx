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
  LayoutGrid,
  Building2,
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

  // Dynamic Branch Breakdown
  const branchEntries = Object.entries(metrics.branchesCount || {});
  const totalSeated = metrics.seatedStudents || roomSeatings.reduce((sum, r) => sum + r.assignedCount, 0) || 235;
  
  const branch1 = branchEntries[0] || ["CSE", Math.round(totalSeated * 0.45)];
  const branch2 = branchEntries[1] || ["ME", Math.round(totalSeated * 0.30)];
  const branch3 = branchEntries[2] || ["ECE", Math.max(0, totalSeated - branch1[1] - branch2[1])];

  const pct1 = Math.round((branch1[1] / totalSeated) * 100) || 45;
  const pct2 = Math.round((branch2[1] / totalSeated) * 100) || 30;
  const pct3 = Math.round((branch3[1] / totalSeated) * 100) || 25;

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
      {/* Title Bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
        <div>
          <div className="flex items-center gap-2 text-xs font-mono font-bold uppercase tracking-wider mb-1 text-slate-700">
            <span className="flex items-center gap-1.5"><Building2 className="h-3.5 w-3.5 text-slate-500" /> {collegeProfile.collegeName}</span>
            <span className="text-slate-400">•</span>
            <span className="bg-[#D4F754] text-black px-2 py-0.5 rounded font-black">{activeSession.title}</span>
          </div>
          <h1 className="text-3xl sm:text-4xl font-black text-slate-900 tracking-tight">
            Examination Overview
          </h1>
          <p className="text-xs text-slate-500 mt-1 font-medium">
            Live multi-branch candidate seating allocations & hall utilization metrics.
          </p>
        </div>

        <div className="flex items-center gap-2 self-start sm:self-auto">
          <span className="text-xs text-slate-500 font-mono font-bold">
            {activeSession.date}
          </span>
          <button
            onClick={() => setIsConfigModalOpen(true)}
            className="px-3.5 py-1.5 rounded-full bg-white border border-slate-200/80 text-xs font-bold text-slate-800 shadow-2xs hover:bg-slate-50 transition"
          >
            Switch Session ⌵
          </button>
        </div>
      </div>

      {/* Main Bento Grid Row */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
        {/* Card 1: Seating Capacity Allocation */}
        <div className="lg:col-span-5 bento-card p-6 flex flex-col justify-between space-y-6">
          {/* Card Top Title & Menu */}
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
              <Zap className="h-4 w-4 text-slate-900 fill-slate-900" />
              <span>Seating Capacity Allocation</span>
            </div>
            <button
              onClick={() => setIsConfigModalOpen(true)}
              className="text-slate-400 hover:text-slate-800"
            >
              <MoreVertical className="h-4 w-4" />
            </button>
          </div>

          {/* Large Metric */}
          <div>
            <div className="flex items-baseline gap-2">
              <span className="text-4xl font-black text-slate-900 font-mono tracking-tight">
                {totalSeated.toLocaleString()}
              </span>
              <span className="px-2 py-0.5 rounded-full bg-[#D4F754] text-black text-[11px] font-black">
                100% Placed
              </span>
            </div>
            <span className="text-xs text-slate-400 font-medium mt-0.5 block">
              candidates seated for {activeSession.timing}
            </span>
          </div>

          {/* Venn / Overlapping Circles Visualization */}
          <div className="relative h-44 w-full flex items-center justify-center my-2">
            {/* Lilac Circle (Branch 1) */}
            <div className="absolute left-6 md:left-12 h-28 w-28 rounded-full bg-[#B8B5FF] flex flex-col items-center justify-center text-slate-900 font-bold shadow-sm z-10 transition-transform hover:scale-105">
              <span className="text-lg font-black leading-none">{branch1[1]}</span>
              <span className="text-[10px] text-slate-800 font-bold uppercase">{branch1[0]}</span>
            </div>

            {/* Dark Charcoal Circle (Branch 2) */}
            <div className="absolute right-6 md:right-12 h-28 w-28 rounded-full bg-[#1E1E22] flex flex-col items-center justify-center text-white font-bold shadow-md z-10 transition-transform hover:scale-105">
              <span className="text-lg font-black leading-none">{branch2[1]}</span>
              <span className="text-[10px] text-slate-300 font-bold uppercase">{branch2[0]}</span>
            </div>

            {/* Neon Lime Small Overlap Circle (Branch 3) */}
            <div className="absolute bottom-1 h-20 w-20 rounded-full bg-[#D4F754] flex flex-col items-center justify-center text-black font-bold shadow-md z-20 border-2 border-white transition-transform hover:scale-105">
              <span className="text-sm font-black leading-none">{branch3[1]}</span>
              <span className="text-[9px] text-black font-black uppercase">{branch3[0]}</span>
            </div>
          </div>

          {/* Progress Percent Breakdown Bars */}
          <div className="space-y-3 pt-2">
            {/* Branch 1 */}
            <div className="flex items-center justify-between text-xs">
              <div className="flex items-center gap-2">
                <span className="font-black text-slate-900">{pct1}%</span>
                <div className="w-28 sm:w-36 h-2 rounded-full bg-slate-100 overflow-hidden">
                  <div className="h-full bg-[#B8B5FF] rounded-full" style={{ width: `${pct1}%` }} />
                </div>
              </div>
              <span className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                {branch1[0]} <span className="h-1.5 w-1.5 rounded-full bg-[#B8B5FF]" />
              </span>
            </div>

            {/* Branch 2 */}
            <div className="flex items-center justify-between text-xs">
              <div className="flex items-center gap-2">
                <span className="font-black text-slate-900">{pct2}%</span>
                <div className="w-28 sm:w-36 h-2 rounded-full bg-slate-100 overflow-hidden">
                  <div className="h-full bg-[#1E1E22] rounded-full" style={{ width: `${pct2}%` }} />
                </div>
              </div>
              <span className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                {branch2[0]} <span className="h-1.5 w-1.5 rounded-full bg-[#1E1E22]" />
              </span>
            </div>

            {/* Branch 3 */}
            <div className="flex items-center justify-between text-xs">
              <div className="flex items-center gap-2">
                <span className="font-black text-slate-900">{pct3}%</span>
                <div className="w-28 sm:w-36 h-2 rounded-full bg-slate-100 overflow-hidden">
                  <div className="h-full bg-[#D4F754] rounded-full" style={{ width: `${pct3}%` }} />
                </div>
              </div>
              <span className="text-[11px] font-bold text-slate-600 flex items-center gap-1">
                {branch3[0]} <span className="h-1.5 w-1.5 rounded-full bg-[#D4F754]" />
              </span>
            </div>
          </div>
        </div>

        {/* Right 7 Cols: 4 Sub-Cards Grid */}
        <div className="lg:col-span-7 grid grid-cols-1 sm:grid-cols-2 gap-5">
          {/* Card 2: Zero Conflict Index */}
          <div className="bento-card p-5 flex flex-col justify-between">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                <ShieldCheck className="h-4 w-4 text-slate-900" />
                <span>Zero Conflict Score</span>
              </div>
              <span className="text-[10px] bg-slate-100 px-2 py-0.5 rounded-full font-mono font-bold text-slate-600">
                Anti-Cheat AI
              </span>
            </div>

            <div className="my-4">
              <div className="flex items-baseline gap-2">
                <span className="text-4xl font-black text-slate-900 font-mono">
                  {metrics.conflictFreeRate || 100}%
                </span>
                <span className="text-xs text-slate-400 font-medium">
                  0 Adjacent Clashes
                </span>
              </div>
              <div className="text-[11px] text-slate-500 mt-1 flex items-center gap-1">
                <span className="h-2 w-2 rounded-full bg-[#D4F754]" />
                <span>Multi-Branch Alternating Matrix</span>
              </div>
            </div>

            <div className="text-[10px] font-mono text-slate-400 pt-2 border-t border-slate-100 flex justify-between font-bold">
              <span>Target 100%</span>
              <span className="text-black">Zero Cheating Risk</span>
            </div>
          </div>

          {/* Card 3: Hall Utilization + Dot Matrix Grid */}
          <div className="bento-card p-5 flex flex-col justify-between">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                <Percent className="h-4 w-4 text-slate-900" />
                <span>Hall Utilization</span>
              </div>
              <span className="text-[10px] bg-[#D4F754] px-2 py-0.5 rounded-full font-mono font-black text-black">
                Optimal
              </span>
            </div>

            <div className="my-2 flex items-center justify-between">
              <div>
                <div className="flex items-baseline gap-1.5">
                  <span className="text-3xl font-black text-slate-900 font-mono">
                    {metrics.utilizationRate || 94}%
                  </span>
                  <span className="px-1.5 py-0.5 rounded-full bg-[#D4F754] text-black text-[10px] font-black">
                    +10%
                  </span>
                </div>
                <span className="text-[11px] text-slate-400 font-medium">Desk Efficiency</span>
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

            <div className="text-[10px] font-mono text-slate-400 pt-2 border-t border-slate-100 flex justify-between font-bold">
              <span>Dynamic Spacing</span>
              <span>No Overcrowding</span>
            </div>
          </div>

          {/* Card 4: Active Examination Halls */}
          <div className="bento-card p-5 flex flex-col justify-between">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                <Activity className="h-4 w-4 text-slate-900" />
                <span>Active Exam Halls</span>
              </div>
              <Link href="/studio" className="text-[11px] font-bold text-black hover:underline">
                View Studio →
              </Link>
            </div>

            <div className="my-3">
              <div className="flex items-baseline gap-2">
                <span className="text-3xl font-black text-slate-900 font-mono">
                  {roomSeatings.length || 3}
                </span>
                <span className="text-xs text-slate-500 font-bold">Halls Live</span>
              </div>
              <div className="text-[11px] text-slate-500 mt-1 font-medium">
                {Math.round(totalSeated / (roomSeatings.length || 1))} Candidates / Hall Average
              </div>
            </div>

            <div className="text-[10px] font-mono text-slate-500 pt-2 border-t border-slate-100 flex justify-between">
              <span className="truncate max-w-[150px]">
                Rooms: {roomSeatings.map((r) => r.roomConfig.roomNumber).join(", ")}
              </span>
              <span className="text-[#D4F754] bg-black px-1.5 py-0.5 rounded text-[9px] font-bold">
                Assigned
              </span>
            </div>
          </div>

          {/* Card 5: Proctors Assigned */}
          <div className="bento-card p-5 flex flex-col justify-between">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-slate-800">
                <GraduationCap className="h-4 w-4 text-slate-900" />
                <span>Invigilators on Duty</span>
              </div>
              <Link href="/invigilators" className="text-[11px] font-bold text-black hover:underline">
                Duty Roster →
              </Link>
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
                <span className="text-[10px] text-slate-400 font-medium">Anti-Bias Rotation</span>
              </div>
            </div>

            <div className="text-[10px] font-mono text-slate-400 pt-2 border-t border-slate-100 flex justify-between font-bold">
              <span>Chief: {collegeProfile.chiefSuperintendent}</span>
            </div>
          </div>

          {/* Card 6: Striped Bar Chart for Real Exam Rooms */}
          <div className="sm:col-span-2 bento-card-dark p-6 flex flex-col justify-between space-y-5">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-xs font-bold text-white">
                <Moon className="h-4 w-4 text-[#D4F754]" />
                <span>Hall Allocation Efficiency & Timelines</span>
              </div>

              <div className="px-3 py-1 rounded-full bg-white/10 text-white text-xs font-bold flex items-center gap-1 border border-white/10">
                <span>{activeSession.title}</span>
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
                    {activeSession.timing.includes("3") ? "3h 00m" : "2h 00m"}
                  </div>
                  <div className="text-[10px] text-slate-400 font-medium mt-0.5">
                    Exam Duration
                  </div>
                </div>
              </div>
            </div>

            {/* Striped Bar Chart (Dynamic from roomSeatings) */}
            <div className="pt-2">
              <div className="flex items-end justify-between gap-3 h-28 px-2">
                {roomSeatings.map((rs, idx) => {
                  const occRate = Math.round((rs.assignedCount / rs.totalCapacity) * 100) || 75;
                  const isHighlight = idx === 0;

                  return (
                    <div key={rs.roomConfig.roomNumber} className="flex-1 flex flex-col items-center gap-2 h-full justify-end">
                      {isHighlight ? (
                        <div className="w-full flex gap-1 h-full items-end">
                          <div
                            className="w-1/2 bg-[#D4F754] rounded-xl shadow-lg shadow-[#D4F754]/20 transition-all duration-300"
                            style={{ height: `${occRate}%` }}
                          />
                          <div
                            className="w-1/2 bg-[#B8B5FF] rounded-xl transition-all duration-300"
                            style={{ height: `${Math.max(25, occRate - 15)}%` }}
                          />
                        </div>
                      ) : (
                        <div
                          className="w-full bg-[#27272A] bg-striped-pattern rounded-xl transition-all duration-300 hover:bg-[#3F3F46]"
                          style={{ height: `${occRate}%` }}
                        />
                      )}
                      <span className={`text-[10px] font-mono ${isHighlight ? "text-[#D4F754] font-bold" : "text-slate-400"}`}>
                        Rm {rs.roomConfig.roomNumber}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Quick Launch Action Ribbon */}
      <div className="bento-card p-5 flex flex-col sm:flex-row items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="h-10 w-10 rounded-2xl bg-[#D4F754] text-black flex items-center justify-center font-bold shadow-xs">
            <LayoutGrid className="h-5 w-5 fill-black stroke-black" />
          </div>
          <div>
            <h3 className="font-black text-sm text-slate-900 tracking-tight">
              Launch Visual 2D Seating Studio
            </h3>
            <p className="text-xs text-slate-500 font-medium">
              Drag-less seat swapping, anti-cheating separation check, and section-wise attendance rosters.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          <Link
            href="/scanner"
            className="px-4 py-2.5 rounded-full bg-white hover:bg-slate-100 text-slate-800 border border-slate-200 text-xs font-bold transition shadow-2xs"
          >
            📷 QR Scanner
          </Link>
          <Link
            href="/studio"
            className="px-5 py-2.5 rounded-full bg-[#161618] hover:bg-black text-white text-xs font-black transition shadow-sm flex items-center gap-1.5"
          >
            <span>Open Studio</span>
            <ArrowUpRight className="h-3.5 w-3.5 text-[#D4F754]" />
          </Link>
        </div>
      </div>
    </div>
  );
}
