"use client";
import React, { useState } from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import {
  UserCheck,
  ShieldAlert,
  Sparkles,
  Clock,
  Plus,
  CheckCircle2,
  AlertCircle,
  FileText,
  UserX,
} from "lucide-react";
import confetti from "canvas-confetti";

interface IncidentLog {
  id: string;
  roomNumber: string;
  studentRoll: string;
  invigilatorName: string;
  type: "Extra Sheet Issued" | "Medical Assistance" | "Disciplinary / Malpractice" | "Late Entry";
  notes: string;
  timestamp: string;
}

export default function InvigilatorRosterPage() {
  const { invigilators, rooms } = useSeating();
  const [activeTab, setActiveTab] = useState<"roster" | "incidents">("roster");

  // Sample incident log state
  const [incidents, setIncidents] = useState<IncidentLog[]>([
    {
      id: "inc-1",
      roomNumber: "302",
      studentRoll: "CSE2026014",
      invigilatorName: "Dr. Rajesh Kulkarni",
      type: "Extra Sheet Issued",
      notes: "Issued Supplementary Booklet #B-8832",
      timestamp: "09:48 AM",
    },
    {
      id: "inc-2",
      roomNumber: "304",
      studentRoll: "ME2026008",
      invigilatorName: "Prof. Sunita Rao",
      type: "Late Entry",
      notes: "Admitted with Dean permission slip (12 min late)",
      timestamp: "09:42 AM",
    },
  ]);

  // Form for new incident
  const [newRoom, setNewRoom] = useState("302");
  const [newRoll, setNewRoll] = useState("");
  const [newType, setNewType] = useState<IncidentLog["type"]>("Extra Sheet Issued");
  const [newNotes, setNewNotes] = useState("");

  const handleAddIncident = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newRoll.trim()) return;

    const newLog: IncidentLog = {
      id: `inc-${Date.now()}`,
      roomNumber: newRoom,
      studentRoll: newRoll.toUpperCase().trim(),
      invigilatorName: "Dr. Rajesh Kulkarni",
      type: newType,
      notes: newNotes || "Action recorded by proctor",
      timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
    };

    setIncidents((prev) => [newLog, ...prev]);
    setNewRoll("");
    setNewNotes("");

    confetti({
      particleCount: 25,
      spread: 50,
      origin: { y: 0.6 },
    });
  };

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8 bg-[#FBF9F4]">
      {/* Header */}
      <div>
        <div className="flex items-center gap-2 text-emerald-700 text-xs font-mono font-semibold uppercase tracking-wider mb-1">
          <Sparkles className="h-4 w-4 text-emerald-600" />
          <span>Faculty Proctoring & Integrity Center</span>
        </div>
        <h1 className="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">
          Invigilator Roster & Incident Logger
        </h1>
        <p className="text-xs text-slate-500 mt-1">
          Automated anti-bias proctor assignment, relief rotation schedule, and live malpractice incident logger.
        </p>
      </div>

      {/* Mode Tabs */}
      <div className="flex items-center gap-2 bg-[#F3EFE6] p-1 rounded-xl border border-[#E2DCCE] w-fit">
        <button
          onClick={() => setActiveTab("roster")}
          className={`px-4 py-1.5 rounded-lg text-xs font-bold transition ${
            activeTab === "roster"
              ? "bg-white text-emerald-900 shadow-sm border border-emerald-100"
              : "text-slate-600 hover:text-slate-900"
          }`}
        >
          👨‍🏫 Proctor Duty Roster
        </button>
        <button
          onClick={() => setActiveTab("incidents")}
          className={`px-4 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1.5 ${
            activeTab === "incidents"
              ? "bg-white text-emerald-900 shadow-sm border border-emerald-100"
              : "text-slate-600 hover:text-slate-900"
          }`}
        >
          <ShieldAlert className="h-3.5 w-3.5 text-amber-600" />
          <span>Incident & Booklet Log</span>
          <span className="text-[10px] px-1.5 py-0.2 rounded-full bg-emerald-100 text-emerald-800 font-semibold">
            {incidents.length}
          </span>
        </button>
      </div>

      {activeTab === "roster" ? (
        /* Invigilators Roster Cards */
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
          {invigilators.map((inv) => (
            <div
              key={inv.id}
              className="bg-white border border-[#E8E2D4] p-5 rounded-2xl shadow-sm space-y-4 hover:border-emerald-300 transition"
            >
              <div className="flex items-center gap-3">
                <img
                  src={inv.image}
                  alt={inv.name}
                  className="h-12 w-12 rounded-xl object-cover border border-emerald-200"
                />
                <div>
                  <h4 className="text-sm font-bold text-slate-900 leading-tight">
                    {inv.name}
                  </h4>
                  <div className="text-xs text-slate-500">{inv.department}</div>
                </div>
              </div>

              <div className="p-3 rounded-xl bg-[#FAF8F3] border border-[#EAE4D6] space-y-1.5 text-xs">
                <div className="flex justify-between">
                  <span className="text-slate-500 font-medium">Assigned Duty:</span>
                  <span className="font-bold font-mono text-emerald-800">
                    {inv.assignedRoom ? `Room ${inv.assignedRoom}` : "Standby Relief"}
                  </span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500 font-medium">Shift Slot:</span>
                  <span className="text-slate-700 font-mono">09:00 - 12:30 PM</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500 font-medium">Anti-Bias Check:</span>
                  <span className="text-emerald-700 font-semibold text-[11px]">
                    ✓ No Dept Conflict
                  </span>
                </div>
              </div>

              <div className="flex items-center justify-between pt-1">
                <span
                  className={`text-[10px] font-bold px-2.5 py-0.5 rounded-full uppercase tracking-wider ${
                    inv.status === "Active"
                      ? "bg-emerald-100 text-emerald-800 border border-emerald-300"
                      : "bg-amber-100 text-amber-800 border border-amber-300"
                  }`}
                >
                  ● {inv.status}
                </span>
                <span className="text-[11px] text-slate-400 font-mono">
                  Duty ID #{inv.id}04
                </span>
              </div>
            </div>
          ))}
        </div>
      ) : (
        /* Incident & Malpractice Logger Mode */
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          {/* Add Incident Form */}
          <div className="lg:col-span-5 bg-white border border-[#E8E2D4] p-6 rounded-3xl shadow-sm space-y-4">
            <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider font-mono flex items-center gap-2">
              <Plus className="h-4 w-4 text-emerald-600" />
              <span>Log Proctor Event / Incident</span>
            </h3>

            <form onSubmit={handleAddIncident} className="space-y-3.5">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Exam Hall
                </label>
                <select
                  value={newRoom}
                  onChange={(e) => setNewRoom(e.target.value)}
                  className="w-full px-3 py-2 rounded-xl bg-[#FAF8F3] border border-[#E2DCCE] text-xs font-semibold text-slate-900 focus:outline-none"
                >
                  {rooms.map((r) => (
                    <option key={r.roomNumber} value={r.roomNumber}>
                      Room {r.roomNumber} ({r.building})
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Student Roll Number *
                </label>
                <input
                  type="text"
                  required
                  placeholder="e.g. CSE2026014"
                  value={newRoll}
                  onChange={(e) => setNewRoll(e.target.value)}
                  className="w-full px-3 py-2 rounded-xl bg-[#FAF8F3] border border-[#E2DCCE] text-xs font-mono text-slate-900 focus:outline-none focus:border-emerald-500"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Event / Incident Type
                </label>
                <select
                  value={newType}
                  onChange={(e) => setNewType(e.target.value as any)}
                  className="w-full px-3 py-2 rounded-xl bg-[#FAF8F3] border border-[#E2DCCE] text-xs font-semibold text-slate-900 focus:outline-none"
                >
                  <option value="Extra Sheet Issued">Extra Sheet / Booklet Issued</option>
                  <option value="Medical Assistance">Medical / Sickness Assistance</option>
                  <option value="Late Entry">Late Entry Candidate</option>
                  <option value="Disciplinary / Malpractice">Disciplinary / Cheating Note</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Proctor Remarks & Booklet Serial
                </label>
                <textarea
                  rows={3}
                  placeholder="Detail serial numbers or reasons..."
                  value={newNotes}
                  onChange={(e) => setNewNotes(e.target.value)}
                  className="w-full px-3 py-2 rounded-xl bg-[#FAF8F3] border border-[#E2DCCE] text-xs text-slate-900 focus:outline-none focus:border-emerald-500"
                />
              </div>

              <button
                type="submit"
                className="w-full py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs uppercase tracking-wider transition shadow-md shadow-emerald-600/20"
              >
                Save Incident Record
              </button>
            </form>
          </div>

          {/* Incident Stream */}
          <div className="lg:col-span-7 bg-white border border-[#E8E2D4] p-6 rounded-3xl shadow-sm space-y-4">
            <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider font-mono">
              Live Hall Event Log ({incidents.length})
            </h3>

            <div className="space-y-3 max-h-[420px] overflow-y-auto no-visible-scrollbar">
              {incidents.map((inc) => (
                <div
                  key={inc.id}
                  className="p-4 rounded-2xl bg-[#FAF8F3] border border-[#E8E2D4] space-y-2 text-xs"
                >
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <span className="font-mono font-bold px-2 py-0.5 rounded bg-white text-slate-900 border border-[#E2DCCE]">
                        Room {inc.roomNumber}
                      </span>
                      <span className="font-mono font-bold text-emerald-800">
                        {inc.studentRoll}
                      </span>
                    </div>

                    <span
                      className={`text-[10px] font-bold px-2.5 py-0.5 rounded-full uppercase tracking-wider ${
                        inc.type === "Extra Sheet Issued"
                          ? "bg-emerald-100 text-emerald-800 border border-emerald-300"
                          : inc.type === "Medical Assistance"
                          ? "bg-blue-100 text-blue-800 border border-blue-300"
                          : inc.type === "Late Entry"
                          ? "bg-amber-100 text-amber-800 border border-amber-300"
                          : "bg-red-100 text-red-800 border border-red-300"
                      }`}
                    >
                      {inc.type}
                    </span>
                  </div>

                  <p className="text-slate-700">{inc.notes}</p>

                  <div className="flex items-center justify-between text-[10px] text-slate-400 font-mono pt-1 border-t border-[#EAE4D6]">
                    <span>Recorded by: {inc.invigilatorName}</span>
                    <span>{inc.timestamp}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
