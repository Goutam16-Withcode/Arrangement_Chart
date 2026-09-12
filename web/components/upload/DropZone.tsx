"use client";
import React, { useState } from "react";
import { UploadCloud, FileSpreadsheet, CheckCircle2 } from "lucide-react";
import { ExcelEngine } from "@/lib/excelEngine";
import { useSeating } from "@/lib/context/SeatingContext";

export const DropZone = () => {
  const { recalculateSeating, setStudentsList } = useSeating();
  const [isDragging, setIsDragging] = useState(false);
  const [uploadedFiles, setUploadedFiles] = useState<{ name: string; type: string; size: string }[]>([]);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);

  const handleFileUpload = async (files: FileList | null) => {
    if (!files || files.length === 0) return;

    const fileList = Array.from(files);
    const newFileSummaries: { name: string; type: string; size: string }[] = [];

    for (const file of fileList) {
      newFileSummaries.push({
        name: file.name,
        type: file.name.endsWith(".xlsx") ? "Excel Workbook" : "Data Sheet",
        size: `${Math.round(file.size / 1024)} KB`,
      });

      if (file.name.toLowerCase().includes("room") || file.name.toLowerCase().includes("excel sheet")) {
        try {
          const rooms = await ExcelEngine.parseRoomLayoutFile(file);
          recalculateSeating(rooms);
          setStatusMessage(`✅ Successfully loaded ${rooms.length} rooms from ${file.name}`);
        } catch (e) {
          console.error(e);
        }
      } else {
        try {
          const branchName = file.name.replace(".xlsx", "").replace("Attendance ", "");
          const students = await ExcelEngine.parseRollNumbersFile(file, branchName);
          setStudentsList(students);
          setStatusMessage(`✅ Successfully loaded ${students.length} students from ${file.name}`);
        } catch (e) {
          console.error(e);
        }
      }
    }

    setUploadedFiles((prev) => [...prev, ...newFileSummaries]);
  };

  return (
    <div className="space-y-4">
      <div
        onDragOver={(e) => {
          e.preventDefault();
          setIsDragging(true);
        }}
        onDragLeave={() => setIsDragging(false)}
        onDrop={(e) => {
          e.preventDefault();
          setIsDragging(false);
          handleFileUpload(e.dataTransfer.files);
        }}
        className={`border-2 border-dashed rounded-3xl p-8 text-center transition-all duration-300 relative overflow-hidden ${
          isDragging
            ? "border-black bg-[#D4F754]/20 scale-[1.01]"
            : "border-slate-300 bg-white hover:border-black shadow-xs"
        }`}
      >
        <input
          type="file"
          id="fileUpload"
          multiple
          accept=".xlsx,.xls,.csv"
          onChange={(e) => handleFileUpload(e.target.files)}
          className="absolute inset-0 opacity-0 cursor-pointer z-10"
        />

        <div className="flex flex-col items-center justify-center space-y-3">
          <div className="p-4 rounded-2xl bg-[#161618] text-[#D4F754] shadow-sm">
            <UploadCloud className="h-8 w-8" />
          </div>

          <div>
            <h3 className="text-base font-black text-slate-900 tracking-tight">
              Drag & Drop your Excel Workbooks here
            </h3>
            <p className="text-xs text-slate-500 mt-1 max-w-sm mx-auto font-medium">
              Supports Room Layouts (`Room Number`, `Rows`, `Benches`) and Roll Number Lists (`F-1`, `S-1`, `T-1`).
            </p>
          </div>

          <div className="text-[11px] font-bold px-4 py-1.5 rounded-full bg-slate-100 text-slate-800 border border-slate-200">
            Click to Browse or Drop .xlsx files
          </div>
        </div>
      </div>

      {/* Status notification */}
      {statusMessage && (
        <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-2xl text-xs text-emerald-900 flex items-center gap-2 font-bold shadow-2xs">
          <CheckCircle2 className="h-4 w-4 flex-shrink-0 text-emerald-600" />
          <span>{statusMessage}</span>
        </div>
      )}

      {/* File summary chips */}
      {uploadedFiles.length > 0 && (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-2">
          {uploadedFiles.map((f, i) => (
            <div
              key={i}
              className="p-3 rounded-2xl bg-white border border-slate-200 flex items-center justify-between text-xs shadow-2xs"
            >
              <div className="flex items-center gap-2.5 truncate">
                <div className="w-7 h-7 rounded-lg bg-slate-100 flex items-center justify-center text-slate-800 flex-shrink-0">
                  <FileSpreadsheet className="h-4 w-4 text-emerald-600" />
                </div>
                <span className="text-slate-900 font-bold truncate">{f.name}</span>
              </div>
              <span className="text-[10px] text-slate-500 font-mono ml-2 font-semibold">{f.size}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
