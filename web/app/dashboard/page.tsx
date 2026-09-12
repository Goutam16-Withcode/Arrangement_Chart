"use client";
import React from "react";
import { BentoGrid, BentoGridItem } from "@/components/ui/bento-grid";
import { AnimatedTooltip } from "@/components/ui/animated-tooltip";
import { useSeating } from "@/lib/context/SeatingContext";
import {
  Users,
  ShieldCheck,
  Building,
  Calendar,
  Layers,
  Sparkles,
  PieChart,
  UserCheck,
  ArrowUpRight,
} from "lucide-react";
import Link from "next/link";

export default function DashboardPage() {
  const { metrics, rooms, invigilators, roomSeatings } = useSeating();

  const invigilatorItems = invigilators.map((inv) => ({
    id: inv.id,
    name: inv.name,
    designation: `${inv.department} • Room ${inv.assignedRoom || "Standby"}`,
    image: inv.image,
    status: inv.status,
  }));

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-10 space-y-8">
      {/* Header Banner */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-6 rounded-3xl glass-panel">
        <div>
          <div className="flex items-center gap-2 text-indigo-400 text-xs font-mono font-semibold uppercase tracking-wider mb-1">
            <Sparkles className="h-4 w-4" />
            <span>Real-Time Exam Control Center</span>
          </div>
          <h1 className="text-2xl sm:text-3xl font-extrabold text-white tracking-tight">
            Institutional Exam Session Overview
          </h1>
          <p className="text-xs text-slate-400 mt-1">
            Active Session: <strong>End-Semester Examinations 2026 (Slot: Morning 09:30 AM)</strong>
          </p>
        </div>

        <div className="flex items-center gap-3">
          <Link
            href="/studio"
            className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white font-semibold text-xs transition flex items-center gap-1.5 shadow-lg shadow-indigo-600/30"
          >
            <span>Open Seating Studio</span>
            <ArrowUpRight className="h-4 w-4" />
          </Link>
        </div>
      </div>

      {/* Bento Grid Architecture */}
      <BentoGrid>
        {/* Item 1: Capacity & Integrity (Span 2) */}
        <BentoGridItem
          className="md:col-span-2"
          title="Overall Exam Hall Capacity & Allocation"
          description="Live allocation rate across all configured blocks with zero-conflict guarantee."
          icon={<Building className="h-5 w-5 text-indigo-400" />}
          header={
            <div className="h-full w-full min-h-[7rem] rounded-xl bg-gradient-to-br from-indigo-950/40 via-slate-900/60 to-purple-950/40 border border-slate-800 p-5 flex flex-col justify-between">
              <div className="grid grid-cols-3 gap-4 text-center">
                <div className="p-3 bg-slate-900/80 rounded-xl border border-slate-800">
                  <div className="text-2xl font-bold font-mono text-white">
                    {metrics.seatedStudents}
                  </div>
                  <div className="text-[10px] text-slate-400 uppercase tracking-wider mt-0.5">
                    Seated Students
                  </div>
                </div>

                <div className="p-3 bg-slate-900/80 rounded-xl border border-slate-800">
                  <div className="text-2xl font-bold font-mono text-indigo-400">
                    {metrics.totalCapacity}
                  </div>
                  <div className="text-[10px] text-slate-400 uppercase tracking-wider mt-0.5">
                    Total Capacity
                  </div>
                </div>

                <div className="p-3 bg-slate-900/80 rounded-xl border border-slate-800">
                  <div className="text-2xl font-bold font-mono text-emerald-400">
                    {metrics.conflictFreeRate}%
                  </div>
                  <div className="text-[10px] text-slate-400 uppercase tracking-wider mt-0.5">
                    Integrity Score
                  </div>
                </div>
              </div>

              {/* Progress bar */}
              <div className="mt-4 space-y-1.5">
                <div className="flex justify-between text-xs font-mono text-slate-400">
                  <span>Occupancy Progress</span>
                  <span>{metrics.utilizationRate}%</span>
                </div>
                <div className="h-2 w-full bg-slate-800 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-gradient-to-r from-indigo-500 to-pink-500 rounded-full transition-all duration-500"
                    style={{ width: `${metrics.utilizationRate}%` }}
                  />
                </div>
              </div>
            </div>
          }
        />

        {/* Item 2: Invigilators On Duty */}
        <BentoGridItem
          className="md:col-span-1"
          title="Invigilators & Proctors"
          description="Assigned faculty coordinators per exam hall with active standby rotation."
          icon={<UserCheck className="h-5 w-5 text-emerald-400" />}
          header={
            <div className="h-full w-full min-h-[7rem] rounded-xl bg-slate-900/80 border border-slate-800 p-5 flex flex-col justify-between">
              <div className="text-xs text-slate-400">Faculty Proctors on Duty:</div>
              <div className="py-2">
                <AnimatedTooltip items={invigilatorItems} />
              </div>
              <div className="text-[11px] text-emerald-400 font-mono">
                ● 3 Assigned • 1 Standby Faculty
              </div>
            </div>
          }
        />

        {/* Item 3: Branch Interleaving Distribution */}
        <BentoGridItem
          className="md:col-span-1"
          title="Interleaved Branches"
          description="Distribution of distinct branch cohorts currently scheduled."
          icon={<PieChart className="h-5 w-5 text-pink-400" />}
          header={
            <div className="h-full w-full min-h-[7rem] rounded-xl bg-slate-900/80 border border-slate-800 p-4 space-y-2">
              {Object.entries(metrics.branchesCount).map(([branch, count], i) => (
                <div
                  key={branch}
                  className="flex items-center justify-between text-xs p-1.5 rounded-lg bg-slate-950/60 border border-slate-800/60"
                >
                  <span className="font-bold text-slate-200">{branch}</span>
                  <span className="font-mono text-indigo-400">{count} students</span>
                </div>
              ))}
            </div>
          }
        />

        {/* Item 4: Room Status Summary (Span 2) */}
        <BentoGridItem
          className="md:col-span-2"
          title="Active Room Roster"
          description="Summary of all active exam halls and individual student capacities."
          icon={<Layers className="h-5 w-5 text-cyan-400" />}
          header={
            <div className="h-full w-full min-h-[7rem] rounded-xl bg-slate-900/80 border border-slate-800 p-4 overflow-y-auto max-h-48 no-visible-scrollbar space-y-2">
              {roomSeatings.map((rs) => (
                <div
                  key={rs.roomConfig.roomNumber}
                  className="flex items-center justify-between p-2.5 rounded-xl bg-slate-950/60 border border-slate-800 text-xs"
                >
                  <div className="flex items-center gap-2">
                    <span className="font-mono font-bold text-white px-2 py-0.5 rounded bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
                      Room {rs.roomConfig.roomNumber}
                    </span>
                    <span className="text-slate-400">
                      {rs.roomConfig.building} (Floor {rs.roomConfig.floor})
                    </span>
                  </div>

                  <div className="flex items-center gap-3">
                    <span className="text-slate-300 font-mono">
                      {rs.assignedCount} / {rs.totalCapacity} Seats
                    </span>
                    <Link
                      href="/studio"
                      className="text-[11px] text-indigo-400 hover:text-indigo-300 underline"
                    >
                      View Map →
                    </Link>
                  </div>
                </div>
              ))}
            </div>
          }
        />
      </BentoGrid>
    </div>
  );
}
