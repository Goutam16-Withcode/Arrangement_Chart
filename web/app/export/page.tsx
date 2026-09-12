"use client";
import React, { useState } from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import { Tabs } from "@/components/ui/tabs";
import { ExcelEngine } from "@/lib/excelEngine";
import {
  Printer,
  FileSpreadsheet,
  Download,
  FileText,
  CreditCard,
  Building2,
  Layers,
  GraduationCap,
} from "lucide-react";
import { QRCodeSVG } from "qrcode.react";

export default function ExportPage() {
  const {
    roomSeatings,
    collegeProfile,
    activeSession,
    setIsConfigModalOpen,
  } = useSeating();
  const [selectedRoom, setSelectedRoom] = useState(roomSeatings[0]?.roomConfig.roomNumber || "302");
  const [sectionGrouping, setSectionGrouping] = useState<"position" | "branch">("position");

  const currentRoom =
    roomSeatings.find((rs) => rs.roomConfig.roomNumber === selectedRoom) ||
    roomSeatings[0];

  const handlePrint = () => {
    window.print();
  };

  const handleExportExcel = () => {
    ExcelEngine.exportSeatingWorkbook(roomSeatings, collegeProfile, activeSession);
  };

  // Group candidates section-wise by seat position (F-1, S-1, T-1)
  const positionSections = ["F-1", "S-1", "T-1", "F-2"].slice(
    0,
    currentRoom?.roomConfig.studentsPerBench || 3
  );

  const getPositionSectionTitle = (pos: string) => {
    switch (pos) {
      case "F-1":
        return "Section 1 • Left Column (F-1)";
      case "S-1":
        return "Section 2 • Middle Column (S-1)";
      case "T-1":
        return "Section 3 • Right Column (T-1)";
      default:
        return `Section (${pos})`;
    }
  };

  // Group candidates section-wise by branch
  const branchSections = Array.from(
    new Set(currentRoom?.attendanceList.map((a) => a.student.branch) || [])
  );

  const tabs = [
    {
      title: "Section-Wise Attendance Register",
      value: "attendance-sheet",
      icon: <FileSpreadsheet className="h-4 w-4" />,
      content: (
        <div className="space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 no-print bg-white p-4 rounded-2xl border border-[#E8E2D4]">
            <div className="flex items-center gap-4 flex-wrap">
              <div className="flex items-center gap-2">
                <span className="text-xs text-slate-500 font-bold">Select Room:</span>
                <select
                  value={selectedRoom}
                  onChange={(e) => setSelectedRoom(e.target.value)}
                  className="px-3.5 py-2 rounded-xl bg-slate-100 border border-slate-300 text-xs text-slate-900 focus:outline-none font-bold"
                >
                  {roomSeatings.map((rs) => (
                    <option key={rs.roomConfig.roomNumber} value={rs.roomConfig.roomNumber}>
                      Room {rs.roomConfig.roomNumber} ({rs.assignedCount} Students)
                    </option>
                  ))}
                </select>
              </div>

              <div className="flex items-center gap-1.5 bg-slate-200/60 p-1 rounded-xl border border-slate-300/60">
                <button
                  onClick={() => setSectionGrouping("position")}
                  className={`px-3 py-1 rounded-lg text-xs font-bold transition ${
                    sectionGrouping === "position"
                      ? "bg-[#161618] text-white shadow-2xs"
                      : "text-slate-600 hover:text-black"
                  }`}
                >
                  By Column (F-1 / S-1 / T-1)
                </button>
                <button
                  onClick={() => setSectionGrouping("branch")}
                  className={`px-3 py-1 rounded-lg text-xs font-bold transition ${
                    sectionGrouping === "branch"
                      ? "bg-[#161618] text-white shadow-2xs"
                      : "text-slate-600 hover:text-black"
                  }`}
                >
                  By Branch / Course Section
                </button>
              </div>
            </div>

            <button
              onClick={handlePrint}
              className="px-4 py-2.5 rounded-full bg-[#161618] hover:bg-black text-white text-xs font-black flex items-center gap-2 shadow-md transition"
            >
              <Printer className="h-4 w-4 text-[#D4F754]" />
              <span>Print Official Attendance Book</span>
            </button>
          </div>

          {/* Printable Section-Wise Attendance Document */}
          {currentRoom && (
            <div className="p-8 bg-white text-slate-900 rounded-[28px] border border-slate-200 shadow-xs space-y-6 max-w-4xl mx-auto font-sans print:p-0 print:border-none print:shadow-none">
              {/* Header */}
              <div className="border-b-2 border-slate-900 pb-3 text-center space-y-1">
                <div className="text-sm font-black uppercase tracking-wide text-slate-900">
                  {collegeProfile.collegeName.toUpperCase()}
                </div>
                <div className="text-xs font-black uppercase tracking-wider text-slate-700 bg-[#D4F754]/30 inline-block px-2 py-0.5 rounded">
                  {activeSession.title.toUpperCase()} • {collegeProfile.academicYear.toUpperCase()}
                </div>
                <div className="text-xs font-semibold text-slate-700">
                  DATE: {activeSession.date} • TIMING: {activeSession.timing} | ROOM {currentRoom.roomConfig.roomNumber} ({currentRoom.assignedCount} Candidates)
                </div>
              </div>

              {/* Renders Section Tables */}
              {sectionGrouping === "position" ? (
                /* Grouped by Seat Position Columns (F-1, S-1, T-1) */
                <div className="space-y-6">
                  {positionSections.map((pos) => {
                    const sectionStudents = currentRoom.attendanceList.filter(
                      (item) => item.position === pos
                    );

                    if (sectionStudents.length === 0) return null;

                    return (
                      <div key={pos} className="space-y-2">
                        {/* Section Header */}
                        <div className="flex items-center justify-between bg-[#FAF8F3] px-3 py-1.5 rounded-lg border border-slate-300">
                          <span className="font-bold text-xs uppercase text-emerald-900 font-mono">
                            {getPositionSectionTitle(pos)}
                          </span>
                          <span className="text-[11px] font-mono text-slate-600">
                            Total: {sectionStudents.length} Candidates
                          </span>
                        </div>

                        {/* Section Table */}
                        <table className="w-full text-left text-xs border-collapse border border-slate-300">
                          <thead>
                            <tr className="bg-slate-100 text-slate-800 font-bold border-b border-slate-300">
                              <th className="p-2 border border-slate-300 w-12 text-center">S.No</th>
                              <th className="p-2 border border-slate-300 w-28">Seat Pos</th>
                              <th className="p-2 border border-slate-300 w-36">Roll Number</th>
                              <th className="p-2 border border-slate-300">Candidate Name</th>
                              <th className="p-2 border border-slate-300 w-28">Department</th>
                              <th className="p-2 border border-slate-300 w-44 text-center">Candidate Signature</th>
                            </tr>
                          </thead>
                          <tbody>
                            {sectionStudents.map((item, idx) => (
                              <tr key={item.serialNo} className="border-b border-slate-200">
                                <td className="p-2 font-mono text-center border border-slate-200">{idx + 1}</td>
                                <td className="p-2 font-mono font-bold border border-slate-200">{item.position}</td>
                                <td className="p-2 font-mono font-extrabold text-slate-900 border border-slate-200">{item.student.rollNo}</td>
                                <td className="p-2 border border-slate-200 font-medium">{item.student.name}</td>
                                <td className="p-2 border border-slate-200">{item.student.branch}</td>
                                <td className="p-2 border border-slate-200 text-center font-mono text-[10px] text-slate-400">
                                  {item.present ? "[ VERIFIED ✓ ]" : "________________"}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    );
                  })}
                </div>
              ) : (
                /* Grouped by Branch / Course Sections */
                <div className="space-y-6">
                  {branchSections.map((branch) => {
                    const sectionStudents = currentRoom.attendanceList.filter(
                      (item) => item.student.branch === branch
                    );

                    return (
                      <div key={branch} className="space-y-2">
                        {/* Section Header */}
                        <div className="flex items-center justify-between bg-[#FAF8F3] px-3 py-1.5 rounded-lg border border-slate-300">
                          <span className="font-bold text-xs uppercase text-emerald-900 font-mono">
                            Section: Department of {branch} (Year {sectionStudents[0]?.student.year || 2})
                          </span>
                          <span className="text-[11px] font-mono text-slate-600">
                            {sectionStudents.length} Candidates Seated
                          </span>
                        </div>

                        {/* Section Table */}
                        <table className="w-full text-left text-xs border-collapse border border-slate-300">
                          <thead>
                            <tr className="bg-slate-100 text-slate-800 font-bold border-b border-slate-300">
                              <th className="p-2 border border-slate-300 w-12 text-center">S.No</th>
                              <th className="p-2 border border-slate-300 w-24">Desk Pos</th>
                              <th className="p-2 border border-slate-300 w-36">Roll Number</th>
                              <th className="p-2 border border-slate-300">Candidate Name</th>
                              <th className="p-2 border border-slate-300 w-32">Subject Code</th>
                              <th className="p-2 border border-slate-300 w-44 text-center">Signature</th>
                            </tr>
                          </thead>
                          <tbody>
                            {sectionStudents.map((item, idx) => (
                              <tr key={item.serialNo} className="border-b border-slate-200">
                                <td className="p-2 font-mono text-center border border-slate-200">{idx + 1}</td>
                                <td className="p-2 font-mono font-bold border border-slate-200">{item.position}</td>
                                <td className="p-2 font-mono font-extrabold text-slate-900 border border-slate-200">{item.student.rollNo}</td>
                                <td className="p-2 border border-slate-200 font-medium">{item.student.name}</td>
                                <td className="p-2 border border-slate-200 font-mono">{item.student.subjectCode}</td>
                                <td className="p-2 border border-slate-200 text-center font-mono text-[10px] text-slate-400">
                                  {item.present ? "[ VERIFIED ✓ ]" : "________________"}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* Proctor Signature Block */}
              <div className="border-t border-slate-300 pt-6 grid grid-cols-2 gap-8 text-xs text-slate-700">
                <div className="space-y-4">
                  <div>Invigilator In-Charge: __________________________</div>
                  <div>Signature: _______________________________</div>
                </div>
                <div className="space-y-4 text-right">
                  <div>Chief Superintendent Seal: {collegeProfile.chiefSuperintendent}</div>
                  <div>Date & Official Timestamp: _______________________</div>
                </div>
              </div>
            </div>
          )}
        </div>
      ),
    },
    {
      title: "Door Notice Chart",
      value: "door-chart",
      icon: <FileText className="h-4 w-4" />,
      content: (
        <div className="space-y-6">
          <div className="flex items-center justify-between no-print">
            <div className="flex items-center gap-2">
              <span className="text-xs text-slate-500 font-semibold">Select Room:</span>
              <select
                value={selectedRoom}
                onChange={(e) => setSelectedRoom(e.target.value)}
                className="px-3 py-1.5 rounded-xl bg-white border border-[#E0D9CB] text-xs text-slate-900 focus:outline-none focus:border-emerald-500 font-semibold"
              >
                {roomSeatings.map((rs) => (
                  <option key={rs.roomConfig.roomNumber} value={rs.roomConfig.roomNumber}>
                    Room {rs.roomConfig.roomNumber} ({rs.assignedCount} Students)
                  </option>
                ))}
              </select>
            </div>

            <button
              onClick={handlePrint}
              className="px-4 py-2.5 rounded-full bg-[#161618] hover:bg-black text-white text-xs font-black flex items-center gap-2 shadow-md transition"
            >
              <Printer className="h-4 w-4 text-[#D4F754]" />
              <span>Print A4 Door Chart</span>
            </button>
          </div>

          {/* Printable Door Notice Document */}
          {currentRoom && (
            <div className="p-8 bg-white text-slate-900 rounded-[28px] border border-slate-200 shadow-xs space-y-6 max-w-4xl mx-auto font-sans print:p-0 print:border-none print:shadow-none">
              {/* Header */}
              <div className="border-b-2 border-slate-900 pb-4 text-center space-y-1">
                <div className="text-sm font-black uppercase tracking-wide text-slate-900">
                  {collegeProfile.collegeName.toUpperCase()}
                </div>
                <div className="text-xs font-black uppercase tracking-wider text-slate-700 bg-[#D4F754]/30 inline-block px-2 py-0.5 rounded">
                  {activeSession.title.toUpperCase()}
                </div>
                <h1 className="text-2xl font-black tracking-tight text-slate-900">
                  EXAMINATION HALL SEATING ARRANGEMENT NOTICE
                </h1>
                <div className="text-sm font-semibold text-slate-700">
                  ROOM {currentRoom.roomConfig.roomNumber} • {currentRoom.roomConfig.building} (Floor {currentRoom.roomConfig.floor}) • Session: {activeSession.timing}
                </div>
              </div>

              {/* Roster Grid */}
              <div className="space-y-4">
                <div className="text-xs font-bold uppercase tracking-wider text-slate-700">
                  Allocated Roll Numbers ({currentRoom.assignedCount} Candidates)
                </div>

                <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-5 gap-2 text-center text-xs">
                  {currentRoom.attendanceList.map((item) => (
                    <div
                      key={item.serialNo}
                      className="p-2.5 border border-slate-200 rounded-xl bg-slate-50 font-mono"
                    >
                      <div className="font-extrabold text-slate-900">{item.student.rollNo}</div>
                      <div className="text-[10px] text-slate-600 font-sans font-medium">
                        Pos: {item.position} • {item.student.branch}
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Footer Notice */}
              <div className="border-t border-slate-200 pt-4 flex items-center justify-between text-[11px] text-slate-500 font-mono font-semibold">
                <span>Total Candidates: {currentRoom.assignedCount}</span>
                <span>Reporting Time: 15 mins prior • Bags & Mobiles Prohibited</span>
              </div>
            </div>
          )}
        </div>
      ),
    },
    {
      title: "Desk QR Stickers",
      value: "desk-stickers",
      icon: <CreditCard className="h-4 w-4" />,
      content: (
        <div className="space-y-6">
          <div className="flex items-center justify-between no-print">
            <p className="text-xs text-slate-500 font-medium">
              Printable adhesive desk slips with individual verification QR codes.
            </p>
            <button
              onClick={handlePrint}
              className="px-4 py-2.5 rounded-full bg-[#161618] hover:bg-black text-white text-xs font-black flex items-center gap-2 shadow-md transition"
            >
              <Printer className="h-4 w-4 text-[#D4F754]" />
              <span>Print Stickers (A4 Label Sheet)</span>
            </button>
          </div>

          {currentRoom && (
            <div className="grid grid-cols-2 md:grid-cols-3 gap-4 p-6 bg-white text-slate-900 rounded-[28px] border border-slate-200 max-w-4xl mx-auto print:p-0 print:border-none">
              {currentRoom.attendanceList.slice(0, 12).map((item) => (
                <div
                  key={item.serialNo}
                  className="p-3.5 border-2 border-dashed border-slate-300 rounded-2xl flex items-center justify-between bg-slate-50"
                >
                  <div className="space-y-0.5">
                    <div className="text-[10px] font-black text-slate-900 font-mono bg-[#D4F754] px-1.5 py-0.5 rounded inline-block">
                      ROOM {currentRoom.roomConfig.roomNumber} • POS: {item.position}
                    </div>
                    <div className="text-sm font-black font-mono text-slate-900">
                      {item.student.rollNo}
                    </div>
                    <div className="text-[11px] text-slate-700 truncate max-w-[120px] font-medium">
                      {item.student.name}
                    </div>
                    <div className="text-[9px] text-slate-500 font-mono">
                      {item.student.branch} • {item.student.subjectCode}
                    </div>
                  </div>

                  <div className="p-1.5 border border-slate-200 rounded-xl bg-white shadow-2xs">
                    <QRCodeSVG
                      value={`VERIFY:${item.student.rollNo}:${currentRoom.roomConfig.roomNumber}`}
                      size={48}
                    />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      ),
    },
    {
      title: "Multi-Sheet Excel (.xlsx)",
      value: "excel-export",
      icon: <Download className="h-4 w-4" />,
      content: (
        <div className="bg-white border border-slate-200 p-8 rounded-[32px] shadow-xs text-center space-y-4 max-w-2xl mx-auto">
          <div className="p-4 rounded-2xl bg-[#161618] text-[#D4F754] w-16 h-16 mx-auto flex items-center justify-center shadow-xs">
            <FileSpreadsheet className="h-8 w-8" />
          </div>

          <h3 className="text-xl font-black text-slate-900 tracking-tight">
            Download Official {activeSession.examMode} Excel Workbook
          </h3>
          <p className="text-xs text-slate-500 max-w-md mx-auto font-medium">
            Exports a formatted institutional workbook with {collegeProfile.collegeName} headers, room seating plans, and section-wise attendance sheets for <strong>{activeSession.title}</strong>.
          </p>

          <button
            onClick={handleExportExcel}
            className="px-6 py-3.5 rounded-full bg-[#161618] hover:bg-black text-white font-black text-xs tracking-wider uppercase transition shadow-md flex items-center gap-2 mx-auto ring-2 ring-[#D4F754]/30"
          >
            <Download className="h-4 w-4 text-[#D4F754]" />
            <span>Download {activeSession.examMode}_Seating_{activeSession.date}.xlsx</span>
          </button>
        </div>
      ),
    },
  ];

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <div className="flex items-center gap-2 text-xs font-mono font-bold uppercase tracking-wider mb-1 flex-wrap text-slate-800">
            <span className="flex items-center gap-1.5"><Building2 className="h-3.5 w-3.5 text-slate-500" /> {collegeProfile.collegeName}</span>
            <span className="text-slate-400">•</span>
            <span>{activeSession.title}</span>
            <span className="px-2 py-0.5 rounded-full bg-[#D4F754] text-black text-[10px] font-bold">
              {activeSession.date}
            </span>
          </div>
          <h1 className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight">
            Export & Print Station
          </h1>
          <p className="text-xs text-slate-500 mt-1">
            Generate section-wise signature registers, door notices, desk stickers, and institutional Excel files.
          </p>
        </div>

        <button
          onClick={() => setIsConfigModalOpen(true)}
          className="px-4 py-2 rounded-full bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 text-xs font-bold flex items-center gap-1.5 transition shadow-2xs self-start md:self-auto"
        >
          <span>⚙️ Edit College & Session Header</span>
        </button>
      </div>

      {/* Tabs Interface */}
      <Tabs tabs={tabs} />
    </div>
  );
}
