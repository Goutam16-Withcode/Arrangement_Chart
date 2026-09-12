"use client";
import React from "react";
import { BentoGrid, BentoGridItem } from "@/components/ui/bento-grid";
import { AnimatedTooltip } from "@/components/ui/animated-tooltip";
import { useSeating } from "@/lib/context/SeatingContext";
import {
  Building,
  Layers,
  Sparkles,
  PieChart,
  UserCheck,
  ArrowUpRight,
} from "lucide-react";
import Link from "next/link";

export default function DashboardPage() {
  const { metrics, invigilators, roomSeatings } = useSeating();

  const invigilatorItems = invigilators.map((inv) => ({
    id: inv.id,
    name: inv.name,
    designation: `${inv.department} • Room ${inv.assignedRoom || "Standby"}`,
    image: inv.image,
    status: inv.status,
  }));

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-10 space-y-8 bg-[#FBF9F4]">
      {/* Header Banner */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-6 rounded-3xl bg-white border border-[#E8E2D4] shadow-sm">
        <div>
          <div className="flex items-center gap-2 text-emerald-700 text-xs font-mono font-semibold uppercase tracking-wider mb-1">
            <Sparkles className="h-4 w-4 text-emerald-600" />
            <span>Real-Time Exam Control Center</span>
          </div>
          <h1 className="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">
            Institutional Exam Session Overview
          </h1>
          <p className="text-xs text-slate-500 mt-1">
            Active Session: <strong>End-Semester Examinations 2026 (Slot: Morning 09:30 AM)</strong>
          </p>
        </div>

        <div className="flex items-center gap-2.5 flex-wrap">
          <Link
            href="/scanner"
            className="px-3.5 py-2 rounded-xl bg-white hover:bg-emerald-50 text-emerald-800 border border-emerald-300 font-bold text-xs transition flex items-center gap-1.5 shadow-2xs"
          >
            <span>📷 QR Scanner</span>
          </Link>
          <Link
            href="/invigilators"
            className="px-3.5 py-2 rounded-xl bg-white hover:bg-[#F6F2E8] text-slate-700 border border-[#E0D9CB] font-semibold text-xs transition flex items-center gap-1.5"
          >
            <span>👨‍🏫 Proctor Roster</span>
          </Link>
          <Link
            href="/3d-twin"
            className="px-3.5 py-2 rounded-xl bg-white hover:bg-[#F6F2E8] text-slate-700 border border-[#E0D9CB] font-semibold text-xs transition flex items-center gap-1.5"
          >
            <span>📦 3D Twin</span>
          </Link>
          <Link
            href="/studio"
            className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs transition flex items-center gap-1.5 shadow-md shadow-emerald-600/20"
          >
            <span>Open Studio</span>
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
          icon={<Building className="h-5 w-5 text-emerald-600" />}
          header={
            <div className="h-full w-full min-h-[7rem] rounded-xl bg-[#FAF8F2] border border-[#EAE4D6] p-5 flex flex-col justify-between">
              <div className="grid grid-cols-3 gap-4 text-center">
                <div className="p-3 bg-white rounded-xl border border-[#E8E2D4] shadow-2xs">
                  <div className="text-2xl font-bold font-mono text-slate-900">
                    {metrics.seatedStudents}
                  </div>
                  <div className="text-[10px] text-slate-500 uppercase tracking-wider mt-0.5 font-semibold">
                    Seated Students
                  </div>
                </div>

                <div className="p-3 bg-white rounded-xl border border-[#E8E2D4] shadow-2xs">
                  <div className="text-2xl font-bold font-mono text-emerald-600">
                    {metrics.totalCapacity}
                  </div>
                  <div className="text-[10px] text-slate-500 uppercase tracking-wider mt-0.5 font-semibold">
                    Total Capacity
                  </div>
                </div>

                <div className="p-3 bg-white rounded-xl border border-[#E8E2D4] shadow-2xs">
                  <div className="text-2xl font-bold font-mono text-teal-600">
                    {metrics.conflictFreeRate}%
                  </div>
                  <div className="text-[10px] text-slate-500 uppercase tracking-wider mt-0.5 font-semibold">
                    Integrity Score
                  </div>
                </div>
              </div>

              {/* Progress bar */}
              <div className="mt-4 space-y-1.5">
                <div className="flex justify-between text-xs font-mono text-slate-600 font-semibold">
                  <span>Occupancy Progress</span>
                  <span className="text-emerald-700">{metrics.utilizationRate}%</span>
                </div>
                <div className="h-2 w-full bg-[#E5DFD1] rounded-full overflow-hidden">
                  <div
                    className="h-full bg-gradient-to-r from-emerald-500 to-teal-500 rounded-full transition-all duration-500"
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
          icon={<UserCheck className="h-5 w-5 text-emerald-600" />}
          header={
            <div className="h-full w-full min-h-[7rem] rounded-xl bg-[#FAF8F2] border border-[#EAE4D6] p-5 flex flex-col justify-between">
              <div className="text-xs text-slate-600 font-medium">Faculty Proctors on Duty:</div>
              <div className="py-2">
                <AnimatedTooltip items={invigilatorItems} />
              </div>
              <div className="text-[11px] text-emerald-800 font-mono font-semibold">
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
          icon={<PieChart className="h-5 w-5 text-teal-600" />}
          header={
            <div className="h-full w-full min-h-[7rem] rounded-xl bg-[#FAF8F2] border border-[#EAE4D6] p-4 space-y-2">
              {Object.entries(metrics.branchesCount).map(([branch, count]) => (
                <div
                  key={branch}
                  className="flex items-center justify-between text-xs p-2 rounded-lg bg-white border border-[#E8E2D4]"
                >
                  <span className="font-bold text-slate-800">{branch}</span>
                  <span className="font-mono text-emerald-700 font-bold">{count} students</span>
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
          icon={<Layers className="h-5 w-5 text-lime-700" />}
          header={
            <div className="h-full w-full min-h-[7rem] rounded-xl bg-[#FAF8F2] border border-[#EAE4D6] p-4 overflow-y-auto max-h-48 no-visible-scrollbar space-y-2">
              {roomSeatings.map((rs) => (
                <div
                  key={rs.roomConfig.roomNumber}
                  className="flex items-center justify-between p-2.5 rounded-xl bg-white border border-[#E8E2D4] text-xs shadow-2xs"
                >
                  <div className="flex items-center gap-2">
                    <span className="font-mono font-bold px-2 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-200">
                      Room {rs.roomConfig.roomNumber}
                    </span>
                    <span className="text-slate-600">
                      {rs.roomConfig.building} (Floor {rs.roomConfig.floor})
                    </span>
                  </div>

                  <div className="flex items-center gap-3">
                    <span className="text-slate-700 font-mono font-semibold">
                      {rs.assignedCount} / {rs.totalCapacity} Seats
                    </span>
                    <Link
                      href="/studio"
                      className="text-[11px] text-emerald-700 font-bold hover:text-emerald-900 underline"
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
