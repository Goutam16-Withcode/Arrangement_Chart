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
  Sparkles,
} from "lucide-react";
import { QRCodeSVG } from "qrcode.react";

export default function ExportPage() {
  const { roomSeatings } = useSeating();
  const [selectedRoom, setSelectedRoom] = useState(roomSeatings[0]?.roomConfig.roomNumber || "302");

  const currentRoom =
    roomSeatings.find((rs) => rs.roomConfig.roomNumber === selectedRoom) ||
    roomSeatings[0];

  const handlePrint = () => {
    window.print();
  };

  const handleExportExcel = () => {
    ExcelEngine.exportSeatingWorkbook(roomSeatings);
  };

  const tabs = [
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
              className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-md shadow-emerald-600/20 transition"
            >
              <Printer className="h-4 w-4" />
              <span>Print A4 Door Chart</span>
            </button>
          </div>

          {/* Printable Door Notice Document */}
          {currentRoom && (
            <div className="p-8 bg-white text-slate-900 rounded-2xl border border-[#E8E2D4] shadow-sm space-y-6 max-w-4xl mx-auto font-sans print:p-0 print:border-none print:shadow-none">
              {/* Header */}
              <div className="border-b-2 border-slate-900 pb-4 text-center space-y-1">
                <div className="text-xs font-bold uppercase tracking-widest text-emerald-800">
                  University Examination Center
                </div>
                <h1 className="text-2xl font-extrabold tracking-tight text-slate-900">
                  HALL SEATING ARRANGEMENT NOTICE
                </h1>
                <div className="text-sm font-semibold text-slate-700">
                  ROOM {currentRoom.roomConfig.roomNumber} • {currentRoom.roomConfig.building} (Floor {currentRoom.roomConfig.floor})
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
                      className="p-2 border border-slate-300 rounded bg-[#FAF8F3] font-mono"
                    >
                      <div className="font-bold text-slate-900">{item.student.rollNo}</div>
                      <div className="text-[10px] text-slate-600 font-sans">
                        Pos: {item.position} • {item.student.branch}
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Footer Notice */}
              <div className="border-t border-slate-300 pt-4 flex items-center justify-between text-[11px] text-slate-500 font-mono">
                <span>Total Candidates: {currentRoom.assignedCount}</span>
                <span>Reporting Time: 09:00 AM • Bags & Mobiles Prohibited</span>
              </div>
            </div>
          )}
        </div>
      ),
    },
    {
      title: "Signature Attendance Sheet",
      value: "attendance-sheet",
      icon: <FileSpreadsheet className="h-4 w-4" />,
      content: (
        <div className="space-y-6">
          <div className="flex items-center justify-between no-print">
            <div className="flex items-center gap-2">
              <span className="text-xs text-slate-500 font-semibold">Select Room:</span>
              <select
                value={selectedRoom}
                onChange={(e) => setSelectedRoom(e.target.value)}
                className="px-3 py-1.5 rounded-xl bg-white border border-[#E0D9CB] text-xs text-slate-900 focus:outline-none font-semibold"
              >
                {roomSeatings.map((rs) => (
                  <option key={rs.roomConfig.roomNumber} value={rs.roomConfig.roomNumber}>
                    Room {rs.roomConfig.roomNumber}
                  </option>
                ))}
              </select>
            </div>

            <button
              onClick={handlePrint}
              className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-md shadow-emerald-600/20 transition"
            >
              <Printer className="h-4 w-4" />
              <span>Print Attendance Roster</span>
            </button>
          </div>

          {/* Printable Attendance Sheet */}
          {currentRoom && (
            <div className="p-8 bg-white text-slate-900 rounded-2xl border border-[#E8E2D4] shadow-sm space-y-4 max-w-4xl mx-auto font-sans print:p-0 print:border-none print:shadow-none">
              <div className="border-b border-slate-800 pb-3 text-center">
                <h2 className="text-lg font-bold">OFFICIAL CANDIDATE ATTENDANCE RECORD</h2>
                <div className="text-xs text-slate-600">
                  Room {currentRoom.roomConfig.roomNumber} | Total Candidates: {currentRoom.assignedCount}
                </div>
              </div>

              <table className="w-full text-left text-xs border-collapse border border-slate-400">
                <thead>
                  <tr className="bg-[#FAF8F3] text-slate-800 font-bold border-b border-slate-400">
                    <th className="p-2 border border-slate-400">#</th>
                    <th className="p-2 border border-slate-400">Seat Tag</th>
                    <th className="p-2 border border-slate-400">Roll Number</th>
                    <th className="p-2 border border-slate-400">Student Name</th>
                    <th className="p-2 border border-slate-400">Branch</th>
                    <th className="p-2 border border-slate-400 w-44">Invigilator Sign</th>
                  </tr>
                </thead>
                <tbody>
                  {currentRoom.attendanceList.map((item) => (
                    <tr key={item.serialNo} className="border-b border-slate-300">
                      <td className="p-2 font-mono border border-slate-300">{item.serialNo}</td>
                      <td className="p-2 font-mono font-bold border border-slate-300">{item.position}</td>
                      <td className="p-2 font-mono font-bold border border-slate-300">{item.student.rollNo}</td>
                      <td className="p-2 border border-slate-300">{item.student.name}</td>
                      <td className="p-2 border border-slate-300">{item.student.branch}</td>
                      <td className="p-2 border border-slate-300 text-slate-400 italic"></td>
                    </tr>
                  ))}
                </tbody>
              </table>
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
            <p className="text-xs text-slate-500">
              Printable adhesive desk slips with individual verification QR codes.
            </p>
            <button
              onClick={handlePrint}
              className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold flex items-center gap-1.5 shadow-md shadow-emerald-600/20 transition"
            >
              <Printer className="h-4 w-4" />
              <span>Print Stickers (A4 Label Sheet)</span>
            </button>
          </div>

          {currentRoom && (
            <div className="grid grid-cols-2 md:grid-cols-3 gap-4 p-6 bg-white text-slate-900 rounded-2xl border border-[#E8E2D4] max-w-4xl mx-auto print:p-0 print:border-none">
              {currentRoom.attendanceList.slice(0, 12).map((item) => (
                <div
                  key={item.serialNo}
                  className="p-3 border-2 border-dashed border-[#D1C9B8] rounded-xl flex items-center justify-between bg-[#FAF8F3]"
                >
                  <div className="space-y-0.5">
                    <div className="text-[10px] font-bold text-emerald-800 font-mono">
                      ROOM {currentRoom.roomConfig.roomNumber} • POS: {item.position}
                    </div>
                    <div className="text-sm font-extrabold font-mono text-slate-900">
                      {item.student.rollNo}
                    </div>
                    <div className="text-[11px] text-slate-700 truncate max-w-[120px]">
                      {item.student.name}
                    </div>
                    <div className="text-[9px] text-slate-500">
                      {item.student.branch} • {item.student.subjectCode}
                    </div>
                  </div>

                  <div className="p-1 border border-[#E0D9CB] rounded bg-white shadow-2xs">
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
        <div className="bg-white border border-[#E8E2D4] p-8 rounded-3xl shadow-sm text-center space-y-4 max-w-2xl mx-auto">
          <div className="p-4 rounded-2xl bg-emerald-50 border border-emerald-200 text-emerald-700 w-16 h-16 mx-auto flex items-center justify-center shadow-2xs">
            <FileSpreadsheet className="h-8 w-8" />
          </div>

          <h3 className="text-xl font-bold text-slate-900">
            Download Institutional Excel Workbook
          </h3>
          <p className="text-xs text-slate-500 max-w-md mx-auto">
            Generates a complete multi-sheet Excel file matching your university layout, containing individual seating charts (`Room 302`) and formatted attendance rosters (`Attendance - Room 302`).
          </p>

          <button
            onClick={handleExportExcel}
            className="px-6 py-3 rounded-2xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs tracking-wider uppercase transition shadow-md shadow-emerald-600/20 flex items-center gap-2 mx-auto"
          >
            <Download className="h-4 w-4" />
            <span>Download SeatingChart_Smart_Output.xlsx</span>
          </button>
        </div>
      ),
    },
  ];

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6 bg-[#FBF9F4]">
      {/* Header */}
      <div>
        <div className="flex items-center gap-2 text-emerald-700 text-xs font-mono font-semibold uppercase tracking-wider mb-1">
          <Sparkles className="h-4 w-4 text-emerald-600" />
          <span>Multi-Format Publishing Hub</span>
        </div>
        <h1 className="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">
          Export & Print Station
        </h1>
        <p className="text-xs text-slate-500 mt-1">
          Generate high-resolution A4 door notices, signature attendance books, desk QR stickers, and Excel workbooks.
        </p>
      </div>

      {/* Tabs Interface */}
      <Tabs tabs={tabs} />
    </div>
  );
}
