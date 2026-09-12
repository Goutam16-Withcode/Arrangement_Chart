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
                ? "bg-[#161618] border-[#161618] text-white shadow-md ring-2 ring-[#D4F754]/30"
                : "bg-white border-[#E5E7EB] text-slate-700 hover:border-black hover:bg-slate-50"
            }`}
          >
            <div
              className={`p-2 rounded-xl transition-all ${
                isSelected
                  ? "bg-[#D4F754] text-black shadow-xs font-bold"
                  : "bg-slate-100 text-slate-600"
              }`}
            >
              <DoorOpen className="h-4 w-4" />
            </div>

            <div>
              <div className={`text-xs font-black tracking-tight ${isSelected ? "text-white" : "text-slate-900"}`}>
                Room {rs.roomConfig.roomNumber}
              </div>
              <div className={`text-[10px] font-mono ${isSelected ? "text-[#D4F754]" : "text-slate-500"}`}>
                {rs.assignedCount}/{rs.totalCapacity} ({occupancyRate}%)
              </div>
            </div>
          </button>
        );
      })}

      <Link
        href="/builder"
        className="flex-shrink-0 px-3.5 py-2.5 rounded-2xl border border-dashed border-slate-300 hover:border-black bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold flex items-center gap-1.5 transition shadow-2xs"
      >
        <Plus className="h-4 w-4 text-black" />
        <span>New Room</span>
      </Link>
    </div>
  );
};
