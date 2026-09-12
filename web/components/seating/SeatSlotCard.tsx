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

  // Branch color styles in Bento Neo-Modern palette
  const getBranchBadgeStyle = (branch?: string) => {
    switch (branch) {
      case "CSE":
        return "bg-[#B8B5FF] text-slate-900 border-[#A5A1F8]";
      case "ME":
        return "bg-[#1E1E22] text-white border-black";
      case "ECE":
        return "bg-[#D4F754] text-black border-[#C2E938]";
      case "IT":
        return "bg-[#B8B5FF] text-slate-900 border-[#A5A1F8]";
      default:
        return "bg-[#D4F754] text-black border-[#C2E938]";
    }
  };

  return (
    <div
      onClick={() => onSelect && onSelect(seat)}
      className={`group relative p-2.5 rounded-xl border transition-all duration-200 cursor-pointer flex flex-col justify-between ${
        isSelected
          ? "border-black ring-2 ring-black bg-slate-100 scale-105 z-20 shadow-md"
          : hasConflict
          ? "border-red-400 bg-red-50 hover:border-red-500"
          : student
          ? "border-slate-200 bg-white hover:border-slate-400 hover:shadow-xs"
          : "border-dashed border-slate-200 bg-slate-50/60 opacity-60 hover:opacity-100"
      }`}
    >
      {/* Top row: Position Tag & Branch Pill */}
      <div className="flex items-center justify-between gap-1 mb-1.5">
        <span className="text-[10px] font-mono font-bold px-1.5 py-0.5 rounded bg-[#F2EDE1] text-slate-700 border border-[#E2DC CE]">
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
        <div className="space-y-0.5">
          <div className="font-mono text-xs font-extrabold text-slate-900 tracking-tight truncate group-hover:text-emerald-700 transition">
            {student.rollNo}
          </div>
          <div className="text-[11px] font-medium text-slate-700 truncate leading-none">
            {student.name}
          </div>
          <div className="text-[9px] text-slate-500 font-mono truncate">
            {student.subjectCode}
          </div>
        </div>
      ) : (
        <div className="py-2 text-center text-[11px] text-slate-400 italic">
          Vacant Seat
        </div>
      )}

      {/* Conflict / Special alert badges */}
      {hasConflict && (
        <div
          title={conflictReason || "Proximity conflict with adjacent seat"}
          className="mt-1.5 flex items-center gap-1 text-[9px] text-red-700 bg-red-100/80 px-1 py-0.5 rounded border border-red-300 font-medium"
        >
          <AlertCircle className="h-3 w-3 flex-shrink-0 text-red-600" />
          <span className="truncate">Conflict</span>
        </div>
      )}

      {student?.hasSpecialNeed && (
        <div className="mt-1 flex items-center gap-1 text-[9px] text-emerald-800 bg-emerald-100 px-1 py-0.5 rounded border border-emerald-300 font-medium">
          <Award className="h-3 w-3 flex-shrink-0 text-emerald-600" />
          <span className="truncate">Priority Desk</span>
        </div>
      )}
    </div>
  );
};
