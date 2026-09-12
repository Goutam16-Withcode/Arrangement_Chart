"use client";
import React, { useState } from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import { RoomSelector } from "@/components/seating/RoomSelector";
import { RoomGrid } from "@/components/seating/RoomGrid";
import { ExcelEngine } from "@/lib/excelEngine";
import {
  Download,
  Printer,
  CheckCircle,
  FileSpreadsheet,
  Users,
  Search,
  Sparkles,
  RefreshCw,
} from "lucide-react";
import Link from "next/link";

export default function StudioPage() {
  const {
    roomSeatings,
    selectedRoomNumber,
    swapSeats,
    toggleAttendance,
    recalculateSeating,
  } = useSeating();

  const [activeTab, setActiveTab] = useState<"visual" | "attendance">("visual");
  const [searchQuery, setSearchQuery] = useState("");

  const currentRoomSeating =
    roomSeatings.find(
      (rs) => rs.roomConfig.roomNumber === selectedRoomNumber
    ) || roomSeatings[0];

  const handleExportExcel = () => {
    ExcelEngine.exportSeatingWorkbook(roomSeatings);
  };

  const filteredAttendance =
    currentRoomSeating?.attendanceList.filter(
      (item) =>
        item.student.rollNo.toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.student.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.student.branch.toLowerCase().includes(searchQuery.toLowerCase())
    ) || [];

  const presentCount =
    currentRoomSeating?.attendanceList.filter((item) => item.present).length || 0;

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
      {/* Studio Header Bar */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <div className="flex items-center gap-2 text-indigo-400 text-xs font-mono font-semibold uppercase tracking-wider mb-1">
            <Sparkles className="h-4 w-4" />
            <span>Interactive 2D Workspace</span>
          </div>
          <h1 className="text-2xl sm:text-3xl font-extrabold text-white tracking-tight">
            Seating Arrangement Studio
          </h1>
          <p className="text-xs text-slate-400 mt-0.5">
            Real-time visual seat allocation, anti-cheating verification, and live attendance logger.
          </p>
        </div>

        {/* Action Buttons */}
        <div className="flex items-center gap-2.5 flex-wrap">
          <button
            onClick={() => recalculateSeating()}
            className="px-3.5 py-2 rounded-xl bg-slate-900 hover:bg-slate-800 text-slate-300 hover:text-white border border-slate-700 text-xs font-semibold flex items-center gap-1.5 transition"
          >
            <RefreshCw className="h-3.5 w-3.5" />
            <span>Re-Optimize</span>
          </button>

          <button
            onClick={handleExportExcel}
            className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold flex items-center gap-1.5 shadow-lg shadow-emerald-600/25 transition"
          >
            <FileSpreadsheet className="h-3.5 w-3.5" />
            <span>Export Excel (.xlsx)</span>
          </button>

          <Link
            href="/export"
            className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold flex items-center gap-1.5 shadow-lg shadow-indigo-600/25 transition"
          >
            <Printer className="h-3.5 w-3.5" />
            <span>Print Station</span>
          </Link>
        </div>
      </div>

      {/* Room Selector Strip */}
      <div className="p-4 rounded-2xl glass-panel">
        <RoomSelector />
      </div>

      {/* Mode Tabs: 2D Classroom Visualizer vs Attendance Sheet */}
      <div className="flex items-center justify-between border-b border-slate-800 pb-3">
        <div className="flex items-center gap-2 bg-slate-900/80 p-1 rounded-xl border border-slate-800">
          <button
            onClick={() => setActiveTab("visual")}
            className={`px-4 py-1.5 rounded-lg text-xs font-semibold transition ${
              activeTab === "visual"
                ? "bg-indigo-600 text-white shadow-md shadow-indigo-600/30"
                : "text-slate-400 hover:text-white"
            }`}
          >
            🖥️ 2D Classroom Grid
          </button>
          <button
            onClick={() => setActiveTab("attendance")}
            className={`px-4 py-1.5 rounded-lg text-xs font-semibold transition flex items-center gap-1.5 ${
              activeTab === "attendance"
                ? "bg-indigo-600 text-white shadow-md shadow-indigo-600/30"
                : "text-slate-400 hover:text-white"
            }`}
          >
            <span>📋 Digital Attendance Roster</span>
            <span className="text-[10px] px-1.5 py-0.2 rounded-full bg-slate-800 text-slate-300">
              {presentCount}/{currentRoomSeating?.assignedCount || 0}
            </span>
          </button>
        </div>

        {activeTab === "attendance" && (
          <div className="relative w-64">
            <Search className="h-3.5 w-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="Filter by roll, name, branch..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-9 pr-3 py-1.5 rounded-xl bg-slate-900 border border-slate-800 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500"
            />
          </div>
        )}
      </div>

      {/* Active Tab View */}
      {currentRoomSeating ? (
        activeTab === "visual" ? (
          <RoomGrid
            roomSeating={currentRoomSeating}
            onSwapSeats={(a, b) => swapSeats(currentRoomSeating.roomConfig.roomNumber, a, b)}
          />
        ) : (
          /* Attendance Sheet Mode */
          <div className="glass-panel p-6 rounded-3xl space-y-4">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-lg font-bold text-white">
                  Attendance Sheet - Room {currentRoomSeating.roomConfig.roomNumber}
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">
                  Click any student row to mark attendance in real time.
                </p>
              </div>

              <div className="flex items-center gap-3 text-xs font-mono">
                <span className="px-3 py-1 rounded-xl bg-emerald-950/60 border border-emerald-800/60 text-emerald-300">
                  Present: {presentCount}
                </span>
                <span className="px-3 py-1 rounded-xl bg-slate-900 border border-slate-800 text-slate-400">
                  Total: {currentRoomSeating.assignedCount}
                </span>
              </div>
            </div>

            <div className="overflow-x-auto rounded-2xl border border-slate-800">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-900/90 text-slate-400 font-mono uppercase text-[10px] tracking-wider border-b border-slate-800">
                  <tr>
                    <th className="py-3 px-4">#</th>
                    <th className="py-3 px-4">Seat Pos</th>
                    <th className="py-3 px-4">Roll Number</th>
                    <th className="py-3 px-4">Student Name</th>
                    <th className="py-3 px-4">Branch</th>
                    <th className="py-3 px-4">Subject</th>
                    <th className="py-3 px-4 text-center">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60">
                  {filteredAttendance.map((item) => (
                    <tr
                      key={item.serialNo}
                      onClick={() =>
                        toggleAttendance(
                          currentRoomSeating.roomConfig.roomNumber,
                          item.serialNo
                        )
                      }
                      className={`cursor-pointer transition ${
                        item.present
                          ? "bg-emerald-950/20 hover:bg-emerald-950/30"
                          : "hover:bg-slate-800/40"
                      }`}
                    >
                      <td className="py-3 px-4 font-mono text-slate-400">{item.serialNo}</td>
                      <td className="py-3 px-4">
                        <span className="font-mono text-[10px] px-2 py-0.5 rounded bg-slate-800 text-slate-300 border border-slate-700">
                          {item.position}
                        </span>
                      </td>
                      <td className="py-3 px-4 font-mono font-bold text-white">
                        {item.student.rollNo}
                      </td>
                      <td className="py-3 px-4 text-slate-200">{item.student.name}</td>
                      <td className="py-3 px-4">
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-indigo-500/10 text-indigo-300 border border-indigo-500/20">
                          {item.student.branch}
                        </span>
                      </td>
                      <td className="py-3 px-4 text-slate-400 font-mono">
                        {item.student.subjectCode}
                      </td>
                      <td className="py-3 px-4 text-center">
                        <span
                          className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider ${
                            item.present
                              ? "bg-emerald-500/20 text-emerald-300 border border-emerald-500/30"
                              : "bg-slate-800 text-slate-400 border border-slate-700"
                          }`}
                        >
                          {item.present ? "Present ✓" : "Absent"}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )
      ) : (
        <div className="p-12 text-center text-slate-500 glass-panel rounded-3xl">
          No room seating available. Please configure rooms or upload files.
        </div>
      )}
    </div>
  );
}
