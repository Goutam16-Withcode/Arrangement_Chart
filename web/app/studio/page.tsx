"use client";
import React, { useState } from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import { RoomSelector } from "@/components/seating/RoomSelector";
import { RoomGrid } from "@/components/seating/RoomGrid";
import { ExcelEngine } from "@/lib/excelEngine";
import {
  Printer,
  FileSpreadsheet,
  Search,
  Sparkles,
  RefreshCw,
  Layers,
} from "lucide-react";
import Link from "next/link";

export default function StudioPage() {
  const {
    roomSeatings,
    selectedRoomNumber,
    swapSeats,
    toggleAttendance,
    recalculateSeating,
    collegeProfile,
    activeSession,
    setIsConfigModalOpen,
  } = useSeating();

  const [activeTab, setActiveTab] = useState<"visual" | "attendance">("visual");
  const [selectedSection, setSelectedSection] = useState<string>("ALL");
  const [searchQuery, setSearchQuery] = useState("");

  const currentRoomSeating =
    roomSeatings.find(
      (rs) => rs.roomConfig.roomNumber === selectedRoomNumber
    ) || roomSeatings[0];

  const handleExportExcel = () => {
    ExcelEngine.exportSeatingWorkbook(roomSeatings, collegeProfile, activeSession);
  };

  const sectionOptions = ["ALL", "F-1", "S-1", "T-1"].slice(
    0,
    (currentRoomSeating?.roomConfig.studentsPerBench || 3) + 1
  );

  const filteredAttendance =
    currentRoomSeating?.attendanceList.filter((item) => {
      const matchesSearch =
        item.student.rollNo.toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.student.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.student.branch.toLowerCase().includes(searchQuery.toLowerCase());

      const matchesSection =
        selectedSection === "ALL" || item.position === selectedSection;

      return matchesSearch && matchesSection;
    }) || [];

  const presentCount =
    currentRoomSeating?.attendanceList.filter((item) => item.present).length || 0;

  return (
    <div className="space-y-6">
      {/* Studio Header Bar */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <div className="flex items-center gap-2 text-xs font-mono font-bold uppercase tracking-wider mb-1 flex-wrap text-slate-800">
            <span>✦ {collegeProfile.collegeName}</span>
            <span className="text-slate-400">•</span>
            <span>{activeSession.title}</span>
            <span className="px-2 py-0.5 rounded-full bg-[#D4F754] text-black text-[10px] font-bold">
              {activeSession.date}
            </span>
          </div>
          <h1 className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight">
            Seating Arrangement Studio
          </h1>
          <p className="text-xs text-slate-500 mt-0.5">
            Real-time visual seat allocation, anti-cheating verification, and live attendance logger.
          </p>
        </div>

        {/* Action Buttons */}
        <div className="flex items-center gap-2.5 flex-wrap">
          <button
            onClick={() => setIsConfigModalOpen(true)}
            className="px-3.5 py-2 rounded-full bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 text-xs font-bold transition shadow-2xs"
          >
            <span>⚙️ Edit College / Session</span>
          </button>

          <button
            onClick={() => recalculateSeating()}
            className="px-3.5 py-2 rounded-full bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 text-xs font-bold flex items-center gap-1.5 transition shadow-2xs"
          >
            <RefreshCw className="h-3.5 w-3.5" />
            <span>Re-Optimize</span>
          </button>

          <button
            onClick={handleExportExcel}
            className="px-4 py-2 rounded-full bg-[#161618] hover:bg-black text-white text-xs font-bold flex items-center gap-1.5 shadow-sm transition"
          >
            <FileSpreadsheet className="h-3.5 w-3.5 text-[#D4F754]" />
            <span>Export Excel</span>
          </button>

          <Link
            href="/export"
            className="px-4 py-2 rounded-full bg-white hover:bg-slate-100 text-slate-900 border border-slate-200 text-xs font-bold flex items-center gap-1.5 shadow-2xs transition"
          >
            <Printer className="h-3.5 w-3.5" />
            <span>Print Station</span>
          </Link>
        </div>
      </div>

      {/* Room Selector Strip */}
      <div className="bento-card p-4">
        <RoomSelector />
      </div>

      {/* Mode Tabs: 2D Classroom Visualizer vs Attendance Sheet */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200 pb-3">
        <div className="flex items-center gap-2 bg-white p-1 rounded-full border border-slate-200 shadow-2xs">
          <button
            onClick={() => setActiveTab("visual")}
            className={`px-4 py-1.5 rounded-full text-xs font-bold transition ${
              activeTab === "visual"
                ? "bg-[#D4F754] text-black font-extrabold shadow-xs"
                : "text-slate-600 hover:text-slate-900"
            }`}
          >
            🖥️ 2D Classroom Grid
          </button>
          <button
            onClick={() => setActiveTab("attendance")}
            className={`px-4 py-1.5 rounded-full text-xs font-bold transition flex items-center gap-1.5 ${
              activeTab === "attendance"
                ? "bg-[#D4F754] text-black font-extrabold shadow-xs"
                : "text-slate-600 hover:text-slate-900"
            }`}
          >
            <span>📋 Digital Attendance Roster</span>
            <span className="text-[10px] px-2 py-0.5 rounded-full bg-[#161618] text-white font-mono font-bold">
              {presentCount}/{currentRoomSeating?.assignedCount || 0}
            </span>
          </button>
        </div>

        {activeTab === "attendance" && (
          <div className="relative w-full sm:w-64">
            <Search className="h-3.5 w-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="Filter by roll, name, branch..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-9 pr-3 py-1.5 rounded-full bg-white border border-slate-200 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-slate-400 shadow-2xs"
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
          <div className="bento-card p-6 space-y-5">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <h3 className="text-lg font-black text-slate-900">
                  Section-Wise Attendance • Room {currentRoomSeating.roomConfig.roomNumber}
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  Click student row to toggle presence. Filter by seat column sections below.
                </p>
              </div>

              {/* Section Selector Pills */}
              <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-full border border-slate-200">
                <span className="text-[11px] font-bold text-slate-500 px-2 font-mono">
                  Section:
                </span>
                {sectionOptions.map((pos) => {
                  const label =
                    pos === "ALL"
                      ? "All Sections"
                      : pos === "F-1"
                      ? "Sec 1 (F-1)"
                      : pos === "S-1"
                      ? "Sec 2 (S-1)"
                      : pos === "T-1"
                      ? "Sec 3 (T-1)"
                      : pos;

                  return (
                    <button
                      key={pos}
                      onClick={() => setSelectedSection(pos)}
                      className={`px-3 py-1 rounded-full text-xs font-bold transition font-mono ${
                        selectedSection === pos
                          ? "bg-[#161618] text-white shadow-2xs"
                          : "text-slate-600 hover:text-black hover:bg-white"
                      }`}
                    >
                      {label}
                    </button>
                  );
                })}
              </div>
            </div>

            <div className="overflow-x-auto rounded-2xl border border-slate-200">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-slate-600 font-mono uppercase text-[10px] tracking-wider border-b border-slate-200">
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
                <tbody className="divide-y divide-slate-100">
                  {filteredAttendance.map((item, index) => (
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
                          ? "bg-slate-50/80 hover:bg-slate-100"
                          : "hover:bg-slate-50/40"
                      }`}
                    >
                      <td className="py-3 px-4 font-mono text-slate-500">{index + 1}</td>
                      <td className="py-3 px-4">
                        <span className="font-mono text-[10px] px-2 py-0.5 rounded bg-slate-100 text-slate-700 border border-slate-200 font-bold">
                          {item.position}
                        </span>
                      </td>
                      <td className="py-3 px-4 font-mono font-black text-slate-900">
                        {item.student.rollNo}
                      </td>
                      <td className="py-3 px-4 text-slate-800 font-medium">{item.student.name}</td>
                      <td className="py-3 px-4">
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-[#D4F754] text-black">
                          {item.student.branch}
                        </span>
                      </td>
                      <td className="py-3 px-4 text-slate-500 font-mono">
                        {item.student.subjectCode}
                      </td>
                      <td className="py-3 px-4 text-center">
                        <span
                          className={`inline-flex items-center gap-1 px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider ${
                            item.present
                              ? "bg-[#161618] text-white shadow-xs"
                              : "bg-slate-100 text-slate-500 border border-slate-200"
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
        <div className="p-12 text-center text-slate-400 bg-white border border-[#E8E2D4] rounded-3xl">
          No room seating available. Please configure rooms or upload files.
        </div>
      )}
    </div>
  );
}
