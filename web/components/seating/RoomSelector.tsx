"use client";
import React from "react";
import { useSeating } from "@/lib/context/SeatingContext";
import { DoorOpen, Plus } from "lucide-react";
import Link from "next/link";

export const RoomSelector = () => {
  const { roomSeatings, selectedRoomNumber, setSelectedRoomNumber } = useSeating();

  return (
    <div className="flex items-center gap-2 overflow-x-auto pb-1 no-visible-scrollbar">
      {roomSeatings.map((rs) => {
        const isSelected = rs.roomConfig.roomNumber === selectedRoomNumber;
        const occupancyRate = Math.round(
          (rs.assignedCount / rs.totalCapacity) * 100
        );

        return (
          <button
            key={rs.roomConfig.roomNumber}
            onClick={() => setSelectedRoomNumber(rs.roomConfig.roomNumber)}
            className={`flex-shrink-0 px-4 py-2.5 rounded-2xl border text-left transition-all duration-200 flex items-center gap-3 ${
              isSelected
                ? "bg-emerald-50 border-emerald-500 text-emerald-950 shadow-sm"
                : "bg-white border-[#E8E2D4] text-slate-700 hover:border-emerald-300 hover:bg-[#FDFBF7]"
            }`}
          >
            <div
              className={`p-2 rounded-xl ${
                isSelected
                  ? "bg-emerald-600 text-white shadow-xs"
                  : "bg-[#F3EFE6] text-slate-600"
              }`}
            >
              <DoorOpen className="h-4 w-4" />
            </div>

            <div>
              <div className="text-xs font-bold font-mono text-slate-900">
                Room {rs.roomConfig.roomNumber}
              </div>
              <div className="text-[10px] text-slate-500">
                {rs.assignedCount}/{rs.totalCapacity} ({occupancyRate}%)
              </div>
            </div>
          </button>
        );
      })}

      <Link
        href="/builder"
        className="flex-shrink-0 px-3.5 py-2.5 rounded-2xl border border-dashed border-[#DDD7C8] hover:border-emerald-500 bg-white hover:bg-emerald-50/50 text-slate-600 hover:text-emerald-800 text-xs font-semibold flex items-center gap-1.5 transition"
      >
        <Plus className="h-4 w-4 text-emerald-600" />
        <span>New Room</span>
      </Link>
    </div>
  );
};
