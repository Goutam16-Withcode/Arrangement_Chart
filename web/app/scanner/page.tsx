"use client";
import React, { useState, useEffect } from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import {
  QrCode,
  Camera,
  CheckCircle2,
  AlertTriangle,
  User,
  Sparkles,
  Clock,
  ShieldCheck,
  RotateCcw,
  Zap,
} from "lucide-react";
import confetti from "canvas-confetti";

interface ScannedRecord {
  rollNo: string;
  name: string;
  branch: string;
  roomNumber: string;
  benchIndex: number;
  positionLabel: string;
  timestamp: string;
  photoUrl: string;
  status: "Verified" | "Invalid Room" | "Already Scanned";
}

export default function ScannerPage() {
  const { roomSeatings, toggleAttendance } = useSeating();
  const [isScanning, setIsScanning] = useState(true);
  const [manualCode, setManualCode] = useState("");
  const [latestScan, setLatestScan] = useState<ScannedRecord | null>(null);
  const [scanHistory, setScanHistory] = useState<ScannedRecord[]>([]);

  // Collect all students in seats
  const allStudentsInSeats: {
    student: any;
    roomNumber: string;
    benchIndex: number;
    positionLabel: string;
    serialNo: number;
  }[] = [];

  roomSeatings.forEach((rs) => {
    rs.seats.forEach((bench) => {
      bench.forEach((seat) => {
        if (seat.student) {
          allStudentsInSeats.push({
            student: seat.student,
            roomNumber: rs.roomConfig.roomNumber,
            benchIndex: seat.benchIndex + 1,
            positionLabel: seat.positionLabel,
            serialNo: rs.attendanceList.find((a) => a.student.id === seat.student?.id)?.serialNo || 1,
          });
        }
      });
    });
  });

  const processScanCode = (code: string) => {
    const clean = code.trim().toUpperCase();
    if (!clean) return;

    // Find student matching the roll number
    const match = allStudentsInSeats.find(
      (item) =>
        clean.includes(item.student.rollNo.toUpperCase()) ||
        clean === item.student.rollNo.toUpperCase()
    );

    const now = new Date().toLocaleTimeString();

    if (match) {
      // Check if already in history
      const alreadyChecked = scanHistory.some(
        (h) => h.rollNo === match.student.rollNo
      );

      const record: ScannedRecord = {
        rollNo: match.student.rollNo,
        name: match.student.name,
        branch: match.student.branch,
        roomNumber: match.roomNumber,
        benchIndex: match.benchIndex,
        positionLabel: match.positionLabel,
        timestamp: now,
        photoUrl: match.student.photoUrl || "https://images.unsplash.com/photo-1534528741775?w=100",
        status: alreadyChecked ? "Already Scanned" : "Verified",
      };

      setLatestScan(record);
      setScanHistory((prev) => [record, ...prev]);

      // Automatically mark attendance in global context
      toggleAttendance(match.roomNumber, match.serialNo);

      if (!alreadyChecked) {
        confetti({
          particleCount: 30,
          spread: 50,
          origin: { y: 0.6 },
          colors: ["#10b981", "#34d399", "#059669"],
        });
      }
    } else {
      setLatestScan({
        rollNo: clean,
        name: "Unknown Candidate",
        branch: "N/A",
        roomNumber: "N/A",
        benchIndex: 0,
        positionLabel: "N/A",
        timestamp: now,
        photoUrl: "https://images.unsplash.com/photo-1534528741775?w=100",
        status: "Invalid Room",
      });
    }

    setManualCode("");
  };

  const simulateQuickScan = () => {
    if (allStudentsInSeats.length === 0) return;
    const randomStudent =
      allStudentsInSeats[Math.floor(Math.random() * allStudentsInSeats.length)];
    processScanCode(
      `EXAM-VERIFY:${randomStudent.student.rollNo}:ROOM-${randomStudent.roomNumber}`
    );
  };

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8 bg-[#FBF9F4]">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <div className="flex items-center gap-2 text-emerald-700 text-xs font-mono font-semibold uppercase tracking-wider mb-1">
            <Sparkles className="h-4 w-4 text-emerald-600" />
            <span>Mobile Invigilator Terminal</span>
          </div>
          <h1 className="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">
            Live QR & Photo Attendance Scanner
          </h1>
          <p className="text-xs text-slate-500 mt-1">
            Scan student desk QR stickers or admit tickets to verify photo ID and register attendance instantly.
          </p>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={simulateQuickScan}
            className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center gap-2 shadow-md shadow-emerald-600/20"
          >
            <Zap className="h-4 w-4" />
            <span>Simulate Live Desk Scan</span>
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
        {/* Scanner Viewport & Manual Input */}
        <div className="lg:col-span-6 space-y-6">
          <div className="bg-white border border-[#E8E2D4] p-6 rounded-3xl shadow-sm space-y-5">
            <div className="flex items-center justify-between">
              <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider font-mono flex items-center gap-2">
                <Camera className="h-4 w-4 text-emerald-600" />
                <span>Optical QR Scanner Viewport</span>
              </h3>
              <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 border border-emerald-300 flex items-center gap-1">
                <span className="h-1.5 w-1.5 rounded-full bg-emerald-500 animate-ping" />
                Camera Ready
              </span>
            </div>

            {/* Simulated Animated Scanner Lens */}
            <div className="relative w-full h-64 rounded-2xl bg-[#1E262B] border-2 border-dashed border-emerald-500/50 flex flex-col items-center justify-center overflow-hidden shadow-inner">
              {/* Animated laser scan line */}
              <div className="absolute inset-x-0 h-1 bg-gradient-to-r from-transparent via-emerald-400 to-transparent animate-[pulseGlow_2s_ease-in-out_infinite] top-1/2 shadow-lg shadow-emerald-400" />

              {/* QR Framing guide */}
              <div className="w-40 h-40 border-2 border-emerald-400/80 rounded-2xl relative flex items-center justify-center">
                <div className="absolute top-0 left-0 w-4 h-4 border-t-2 border-l-2 border-emerald-300 -translate-x-1 -translate-y-1" />
                <div className="absolute top-0 right-0 w-4 h-4 border-t-2 border-r-2 border-emerald-300 translate-x-1 -translate-y-1" />
                <div className="absolute bottom-0 left-0 w-4 h-4 border-b-2 border-l-2 border-emerald-300 -translate-x-1 translate-y-1" />
                <div className="absolute bottom-0 right-0 w-4 h-4 border-b-2 border-r-2 border-emerald-300 translate-x-1 translate-y-1" />
                <QrCode className="h-16 w-16 text-emerald-400/40 animate-pulse" />
              </div>

              <div className="text-[11px] text-emerald-200/80 font-mono mt-3">
                Align QR Code within the frame
              </div>
            </div>

            {/* Manual Roll Number / Barcode Entry */}
            <form
              onSubmit={(e) => {
                e.preventDefault();
                processScanCode(manualCode);
              }}
              className="flex items-center gap-2 pt-2"
            >
              <input
                type="text"
                placeholder="Or type/paste Roll Code (e.g. CSE2026001)..."
                value={manualCode}
                onChange={(e) => setManualCode(e.target.value)}
                className="flex-1 px-4 py-2.5 rounded-xl bg-[#FAF8F3] border border-[#E2DCCE] text-xs font-mono text-slate-900 placeholder-slate-400 focus:outline-none focus:border-emerald-500"
              />
              <button
                type="submit"
                className="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs uppercase tracking-wider transition"
              >
                Verify
              </button>
            </form>
          </div>

          {/* Quick Stats Bar */}
          <div className="grid grid-cols-3 gap-3">
            <div className="p-4 rounded-2xl bg-white border border-[#E8E2D4] text-center shadow-2xs">
              <div className="text-xl font-bold font-mono text-emerald-600">
                {scanHistory.length}
              </div>
              <div className="text-[10px] text-slate-500 font-semibold uppercase mt-0.5">
                Total Scanned
              </div>
            </div>

            <div className="p-4 rounded-2xl bg-white border border-[#E8E2D4] text-center shadow-2xs">
              <div className="text-xl font-bold font-mono text-slate-900">
                {scanHistory.filter((s) => s.status === "Verified").length}
              </div>
              <div className="text-[10px] text-slate-500 font-semibold uppercase mt-0.5">
                Verified In
              </div>
            </div>

            <div className="p-4 rounded-2xl bg-white border border-[#E8E2D4] text-center shadow-2xs">
              <div className="text-xl font-bold font-mono text-amber-600">
                {scanHistory.filter((s) => s.status === "Invalid Room").length}
              </div>
              <div className="text-[10px] text-slate-500 font-semibold uppercase mt-0.5">
                Alerts / Mismatch
              </div>
            </div>
          </div>
        </div>

        {/* Verification Card & Live Check-in Stream */}
        <div className="lg:col-span-6 space-y-6">
          {/* Active Candidate Photo Match Card */}
          {latestScan ? (
            <div className="bg-white border border-[#E8E2D4] p-6 rounded-3xl shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-[#E8E2D4] pb-3">
                <span className="text-xs font-mono font-bold text-slate-500 uppercase tracking-wider">
                  Biometric & Photo ID Match
                </span>
                <span
                  className={`text-[10px] font-bold px-2.5 py-0.5 rounded-full uppercase tracking-wider ${
                    latestScan.status === "Verified"
                      ? "bg-emerald-100 text-emerald-800 border border-emerald-300"
                      : latestScan.status === "Already Scanned"
                      ? "bg-amber-100 text-amber-800 border border-amber-300"
                      : "bg-red-100 text-red-800 border border-red-300"
                  }`}
                >
                  {latestScan.status}
                </span>
              </div>

              <div className="flex items-center gap-4">
                <div className="h-20 w-20 rounded-2xl border-2 border-emerald-400 bg-slate-100 overflow-hidden shadow-sm flex-shrink-0">
                  <img
                    src={latestScan.photoUrl}
                    alt={latestScan.name}
                    className="h-full w-full object-cover"
                  />
                </div>

                <div className="space-y-1">
                  <h3 className="text-lg font-bold text-slate-900 leading-tight">
                    {latestScan.name}
                  </h3>
                  <div className="text-xs font-mono font-bold text-emerald-700">
                    {latestScan.rollNo}
                  </div>
                  <div className="text-xs text-slate-600">
                    Department: <strong>{latestScan.branch}</strong>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2 pt-2 text-xs">
                <div className="p-2.5 rounded-xl bg-[#FAF8F3] border border-[#E8E2D4]">
                  <span className="text-slate-500 block text-[10px] uppercase font-mono">
                    Assigned Room
                  </span>
                  <span className="font-bold text-slate-800 font-mono">
                    Room {latestScan.roomNumber}
                  </span>
                </div>

                <div className="p-2.5 rounded-xl bg-[#FAF8F3] border border-[#E8E2D4]">
                  <span className="text-slate-500 block text-[10px] uppercase font-mono">
                    Desk Coordinate
                  </span>
                  <span className="font-bold text-emerald-800 font-mono">
                    Bench {latestScan.benchIndex} ({latestScan.positionLabel})
                  </span>
                </div>
              </div>

              <div className="p-2.5 rounded-xl bg-emerald-50 border border-emerald-200 text-xs text-emerald-800 flex items-center justify-between font-semibold">
                <div className="flex items-center gap-1.5">
                  <CheckCircle2 className="h-4 w-4 text-emerald-600" />
                  <span>Attendance Timestamp Recorded</span>
                </div>
                <span className="font-mono text-[10px]">{latestScan.timestamp}</span>
              </div>
            </div>
          ) : (
            <div className="bg-white border border-[#E8E2D4] p-10 rounded-3xl shadow-sm text-center space-y-2">
              <QrCode className="h-10 w-10 text-emerald-600 mx-auto opacity-40" />
              <h4 className="text-sm font-bold text-slate-900">Awaiting First Scan</h4>
              <p className="text-xs text-slate-500">
                Click &quot;Simulate Live Desk Scan&quot; or enter a roll number to inspect verification.
              </p>
            </div>
          )}

          {/* Live Check-in Stream */}
          <div className="bg-white border border-[#E8E2D4] p-6 rounded-3xl shadow-sm space-y-3">
            <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wider font-mono">
              Live Invigilator Scan Stream
            </h4>

            <div className="space-y-2 max-h-56 overflow-y-auto no-visible-scrollbar">
              {scanHistory.length > 0 ? (
                scanHistory.map((item, idx) => (
                  <div
                    key={idx}
                    className="p-2.5 rounded-xl bg-[#FAF8F3] border border-[#E8E2D4] flex items-center justify-between text-xs"
                  >
                    <div className="flex items-center gap-2">
                      <CheckCircle2 className="h-3.5 w-3.5 text-emerald-600 flex-shrink-0" />
                      <span className="font-bold font-mono text-slate-900">
                        {item.rollNo}
                      </span>
                      <span className="text-slate-600 truncate max-w-[120px]">
                        {item.name}
                      </span>
                    </div>

                    <div className="flex items-center gap-2 font-mono text-[10px] text-slate-500">
                      <span>Room {item.roomNumber}</span>
                      <span>•</span>
                      <span>{item.timestamp}</span>
                    </div>
                  </div>
                ))
              ) : (
                <div className="py-4 text-center text-xs text-slate-400 italic">
                  No scan events recorded in this session yet.
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
