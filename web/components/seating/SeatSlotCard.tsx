"use client";
import React from "react";
import { SeatSlot } from "@/lib/types";
import { AlertCircle, User, Award, CheckCircle2 } from "lucide-react";

interface SeatSlotCardProps {
  seat: SeatSlot;
  isSelected?: boolean;
  onSelect?: (seat: SeatSlot) => void;
}

export const SeatSlotCard: React.FC<SeatSlotCardProps> = ({
  seat,
  isSelected = false,
  onSelect,
}) => {
  const { student, positionLabel, hasConflict, conflictReason } = seat;

  // Branch color styles
  const getBranchBadgeStyle = (branch?: string) => {
    switch (branch) {
      case "CSE":
        return "bg-indigo-500/20 text-indigo-300 border-indigo-500/40";
      case "ME":
        return "bg-emerald-500/20 text-emerald-300 border-emerald-500/40";
      case "ECE":
        return "bg-amber-500/20 text-amber-300 border-amber-500/40";
      case "IT":
        return "bg-cyan-500/20 text-cyan-300 border-cyan-500/40";
      default:
        return "bg-purple-500/20 text-purple-300 border-purple-500/40";
    }
  };

  return (
    <div
      onClick={() => onSelect && onSelect(seat)}
      className={`group relative p-2.5 rounded-xl border transition-all duration-200 cursor-pointer flex flex-col justify-between ${
        isSelected
          ? "border-pink-500 ring-2 ring-pink-500/50 bg-pink-500/10 scale-105 z-20 shadow-lg shadow-pink-500/20"
          : hasConflict
          ? "border-red-500/60 bg-red-950/30 hover:border-red-400"
          : student
          ? "border-slate-800 bg-slate-900/80 hover:border-indigo-500/60 hover:bg-slate-800/80"
          : "border-dashed border-slate-800 bg-slate-950/40 opacity-50 hover:opacity-80"
      }`}
    >
      {/* Top row: Position Tag & Branch Pill */}
      <div className="flex items-center justify-between gap-1 mb-1.5">
        <span className="text-[10px] font-mono font-bold px-1.5 py-0.5 rounded bg-slate-800 text-slate-300 border border-slate-700">
          {positionLabel}
        </span>

        {student && (
          <span
            className={`text-[9px] font-bold px-1.5 py-0.5 rounded border uppercase tracking-wider ${getBranchBadgeStyle(
              student.branch
            )}`}
          >
            {student.branch}
          </span>
        )}
      </div>

      {/* Main Seat Content */}
      {student ? (
        <div className="space-y-1">
          <div className="font-mono text-xs font-bold text-white tracking-tight truncate group-hover:text-indigo-300 transition">
            {student.rollNo}
          </div>
          <div className="text-[11px] text-slate-400 truncate leading-none">
            {student.name}
          </div>
          <div className="text-[9px] text-slate-500 font-mono truncate">
            {student.subjectCode}
          </div>
        </div>
      ) : (
        <div className="py-2 text-center text-[11px] text-slate-500 italic">
          Vacant Seat
        </div>
      )}

      {/* Conflict / Special alert badges */}
      {hasConflict && (
        <div
          title={conflictReason || "Proximity conflict with adjacent seat"}
          className="mt-1.5 flex items-center gap-1 text-[9px] text-red-400 bg-red-950/60 px-1 py-0.5 rounded border border-red-800/60"
        >
          <AlertCircle className="h-3 w-3 flex-shrink-0" />
          <span className="truncate">Conflict</span>
        </div>
      )}

      {student?.hasSpecialNeed && (
        <div className="mt-1 flex items-center gap-1 text-[9px] text-cyan-300 bg-cyan-950/40 px-1 py-0.5 rounded border border-cyan-800/40">
          <Award className="h-3 w-3 flex-shrink-0" />
          <span className="truncate">Front Row Priority</span>
        </div>
      )}
    </div>
  );
};
