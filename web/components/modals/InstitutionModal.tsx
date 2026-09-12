"use client";
import React, { useState, useEffect } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { useSeating } from "@/lib/context/SeatingContext";
import {
  Building2,
  Calendar,
  Clock,
  GraduationCap,
  Sparkles,
  X,
  Check,
  RotateCcw,
  Layers,
  FileCheck2,
  Sliders,
  Shield,
  Zap,
} from "lucide-react";
import confetti from "canvas-confetti";

export const InstitutionModal = () => {
  const {
    collegeProfile,
    updateCollegeProfile,
    activeSession,
    updateActiveSession,
    sessions,
    switchSession,
    isConfigModalOpen,
    setIsConfigModalOpen,
  } = useSeating();

  const [activeTab, setActiveTab] = useState<"profile" | "session" | "presets">("profile");

  // Local form state for Profile
  const [name, setName] = useState(collegeProfile.collegeName);
  const [code, setCode] = useState(collegeProfile.collegeCode);
  const [academicYear, setAcademicYear] = useState(collegeProfile.academicYear);
  const [semester, setSemester] = useState(collegeProfile.semester);
  const [centerCode, setCenterCode] = useState(collegeProfile.examCenterCode);
  const [superintendent, setSuperintendent] = useState(collegeProfile.chiefSuperintendent);

  // Local form state for Session
  const [sessionTitle, setSessionTitle] = useState(activeSession.title);
  const [sessionDate, setSessionDate] = useState(activeSession.date);
  const [sessionTiming, setSessionTiming] = useState(activeSession.timing);
  const [sessionSlot, setSessionSlot] = useState(activeSession.slot);

  // Sync state whenever collegeProfile or activeSession changes
  useEffect(() => {
    setName(collegeProfile.collegeName);
    setCode(collegeProfile.collegeCode);
    setAcademicYear(collegeProfile.academicYear);
    setSemester(collegeProfile.semester);
    setCenterCode(collegeProfile.examCenterCode);
    setSuperintendent(collegeProfile.chiefSuperintendent);
  }, [collegeProfile]);

  useEffect(() => {
    setSessionTitle(activeSession.title);
    setSessionDate(activeSession.date);
    setSessionTiming(activeSession.timing);
    setSessionSlot(activeSession.slot);
  }, [activeSession]);

  const handleSave = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    updateCollegeProfile({
      collegeName: name,
      collegeCode: code,
      academicYear,
      semester,
      examCenterCode: centerCode,
      chiefSuperintendent: superintendent,
    });
    updateActiveSession({
      title: sessionTitle,
      date: sessionDate,
      timing: sessionTiming,
      slot: sessionSlot,
    });

    confetti({
      particleCount: 30,
      spread: 60,
      origin: { y: 0.5 },
    });

    setIsConfigModalOpen(false);
  };

  const presetInstitutions = [
    {
      label: "Autonomous Engineering Institute",
      badge: "Tech / Engg",
      name: "Apex Institute of Science & Technology",
      code: "AIST-2026",
      center: "MAIN-BLOCK-A",
      year: "Academic Session 2025–2026",
      sem: "End-Semester Major Theory Exam",
      superintendent: "Prof. Controller of Examinations",
      examTitle: "End-Semester Major Theory Examination • Slot 1",
      date: "Day 1 (Morning Session)",
      timing: "09:30 AM - 12:30 PM (3.0 Hours)",
      slot: "Slot 1 (Morning)",
    },
    {
      label: "National Central University",
      badge: "University",
      name: "National Autonomous University",
      code: "NAU-1001",
      center: "CONVOCATION-HALL",
      year: "Academic Year 2025–2026",
      sem: "Even Semester Major Board Examination",
      superintendent: "Dr. Registrar & Chief Superintendent",
      examTitle: "University Major Theory Examination • Day 1",
      date: "Day 1 (Morning Session)",
      timing: "09:30 AM - 12:30 PM (3.0 Hours)",
      slot: "Slot 1",
    },
    {
      label: "Mid-Semester Assessment (MST)",
      badge: "MST Test",
      name: "University School of Engineering",
      code: "USE-404",
      center: "ACADEMIC-WING-3",
      year: "Academic Session 2025–2026",
      sem: "Continuous Internal Assessment (MST-1)",
      superintendent: "Dean of Academic Affairs",
      examTitle: "Mid-Semester Assessment Test (MST) • Slot 1",
      date: "Day 1 (Morning Session)",
      timing: "10:00 AM - 11:30 AM (1.5 Hours)",
      slot: "Slot 1 (Morning)",
    },
    {
      label: "Global Business & Management School",
      badge: "Management",
      name: "Global Institute of Management & Research",
      code: "GIMR-808",
      center: "EXECUTIVE-HALL",
      year: "Trimester Session 2025–2026",
      sem: "Trimester Final Evaluation",
      superintendent: "Director & Examination Head",
      examTitle: "Trimester Comprehensive Assessment • Slot 1",
      date: "Day 1 (Morning Session)",
      timing: "09:30 AM - 12:30 PM (3.0 Hours)",
      slot: "Slot 1 (Morning)",
    },
    {
      label: "Medical & Health Sciences College",
      badge: "Health / Medical",
      name: "Metropolitan College of Health & Pharmacy",
      code: "MCHP-302",
      center: "CLINICAL-AUDITORIUM",
      year: "Professional Batch 2025–2026",
      sem: "Annual Professional Board Exam",
      superintendent: "Medical Superintendent & Board Chair",
      examTitle: "Professional Theory Examination • Slot 1",
      date: "Day 1 (Morning Session)",
      timing: "09:30 AM - 12:30 PM (3.0 Hours)",
      slot: "Slot 1 (Morning)",
    },
  ];

  const applyPreset = (preset: typeof presetInstitutions[0]) => {
    setName(preset.name);
    setCode(preset.code);
    setCenterCode(preset.center);
    setAcademicYear(preset.year);
    setSemester(preset.sem);
    setSuperintendent(preset.superintendent);
    setSessionTitle(preset.examTitle);
    setSessionDate(preset.date);
    setSessionTiming(preset.timing);
    setSessionSlot(preset.slot);

    updateCollegeProfile({
      collegeName: preset.name,
      collegeCode: preset.code,
      examCenterCode: preset.center,
      academicYear: preset.year,
      semester: preset.sem,
      chiefSuperintendent: preset.superintendent,
    });

    updateActiveSession({
      title: preset.examTitle,
      date: preset.date,
      timing: preset.timing,
      slot: preset.slot,
    });

    confetti({
      particleCount: 25,
      spread: 50,
      origin: { y: 0.5 },
    });
  };

  if (!isConfigModalOpen) return null;

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
        <motion.div
          initial={{ opacity: 0, scale: 0.95, y: 15 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.95, y: 15 }}
          transition={{ type: "spring", stiffness: 350, damping: 28 }}
          className="bg-white rounded-3xl border border-slate-200 shadow-2xl w-full max-w-3xl overflow-hidden max-h-[92vh] flex flex-col"
        >
          {/* Modal Header */}
          <div className="p-5 sm:p-6 bg-[#161618] text-white flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="h-10 w-10 rounded-2xl bg-[#D4F754] text-black flex items-center justify-center font-bold shadow-sm">
                <Building2 className="h-5 w-5" />
              </div>
              <div>
                <h2 className="text-lg sm:text-xl font-black tracking-tight flex items-center gap-2">
                  <span>Institution & Exam Session Settings</span>
                  <span className="px-2.5 py-0.5 rounded-full text-[10px] font-mono font-bold bg-[#D4F754] text-black">
                    Customizer
                  </span>
                </h2>
                <p className="text-xs text-slate-400 mt-0.5 font-medium">
                  Automatically updates seating charts, attendance sheets, and passes across the suite.
                </p>
              </div>
            </div>

            <button
              onClick={() => setIsConfigModalOpen(false)}
              className="p-2 rounded-full text-slate-400 hover:text-white hover:bg-white/10 transition"
            >
              <X className="h-5 w-5" />
            </button>
          </div>

          {/* Navigation Tabs */}
          <div className="flex border-b border-slate-200 bg-slate-100 px-6 gap-2 pt-2">
            <button
              onClick={() => setActiveTab("profile")}
              className={`pb-3 px-3 text-xs font-bold transition flex items-center gap-2 border-b-2 ${
                activeTab === "profile"
                  ? "border-black text-black font-extrabold"
                  : "border-transparent text-slate-500 hover:text-slate-800"
              }`}
            >
              <GraduationCap className="h-4 w-4" />
              <span>College Profile</span>
            </button>

            <button
              onClick={() => setActiveTab("session")}
              className={`pb-3 px-3 text-xs font-bold transition flex items-center gap-2 border-b-2 ${
                activeTab === "session"
                  ? "border-black text-black font-extrabold"
                  : "border-transparent text-slate-500 hover:text-slate-800"
              }`}
            >
              <Calendar className="h-4 w-4" />
              <span>Exam Session</span>
            </button>

            <button
              onClick={() => setActiveTab("presets")}
              className={`pb-3 px-3 text-xs font-bold transition flex items-center gap-2 border-b-2 ${
                activeTab === "presets"
                  ? "border-black text-black font-extrabold"
                  : "border-transparent text-slate-500 hover:text-slate-800"
              }`}
            >
              <Sparkles className="h-4 w-4 text-black" />
              <span>1-Click Presets</span>
            </button>
          </div>

          {/* Modal Body */}
          <div className="p-6 overflow-y-auto space-y-6 flex-1 bg-[#F2F3F5]">
            {/* Tab 1: College Profile */}
            {activeTab === "profile" && (
              <div className="space-y-4">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="sm:col-span-2 space-y-1.5">
                    <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                      <Building2 className="h-3.5 w-3.5 text-emerald-600" />
                      <span>University / Institution / College Name</span>
                    </label>
                    <input
                      type="text"
                      value={name}
                      onChange={(e) => setName(e.target.value)}
                      placeholder="e.g. Apex University of Science & Technology"
                      className="w-full px-3.5 py-2.5 rounded-xl bg-white border border-[#E0D9CB] text-xs font-semibold text-slate-900 focus:outline-none focus:border-emerald-500 transition shadow-2xs"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-700">
                      College / Institution Code
                    </label>
                    <input
                      type="text"
                      value={code}
                      onChange={(e) => setCode(e.target.value)}
                      placeholder="e.g. AUTM-2026 or UNIV-101"
                      className="w-full px-3.5 py-2.5 rounded-xl bg-white border border-[#E0D9CB] text-xs font-mono font-semibold text-slate-900 focus:outline-none focus:border-emerald-500 transition shadow-2xs"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-700">
                      Exam Center / Hall Code
                    </label>
                    <input
                      type="text"
                      value={centerCode}
                      onChange={(e) => setCenterCode(e.target.value)}
                      placeholder="e.g. EXAM-CTR-01 or MAIN-WING"
                      className="w-full px-3.5 py-2.5 rounded-xl bg-white border border-[#E0D9CB] text-xs font-mono font-semibold text-slate-900 focus:outline-none focus:border-emerald-500 transition shadow-2xs"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-700">
                      Academic Year / Session
                    </label>
                    <input
                      type="text"
                      value={academicYear}
                      onChange={(e) => setAcademicYear(e.target.value)}
                      placeholder="e.g. Academic Session 2025–2026"
                      className="w-full px-3.5 py-2.5 rounded-xl bg-white border border-[#E0D9CB] text-xs font-semibold text-slate-900 focus:outline-none focus:border-emerald-500 transition shadow-2xs"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-700">
                      Semester / Examination Term
                    </label>
                    <input
                      type="text"
                      value={semester}
                      onChange={(e) => setSemester(e.target.value)}
                      placeholder="e.g. End-Semester Major Examination"
                      className="w-full px-3.5 py-2.5 rounded-xl bg-white border border-[#E0D9CB] text-xs font-semibold text-slate-900 focus:outline-none focus:border-emerald-500 transition shadow-2xs"
                    />
                  </div>

                  <div className="sm:col-span-2 space-y-1.5">
                    <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                      <Shield className="h-3.5 w-3.5 text-emerald-600" />
                      <span>Controller of Examinations / Chief Superintendent</span>
                    </label>
                    <input
                      type="text"
                      value={superintendent}
                      onChange={(e) => setSuperintendent(e.target.value)}
                      placeholder="e.g. Dr. Controller of Examinations"
                      className="w-full px-3.5 py-2.5 rounded-xl bg-white border border-[#E0D9CB] text-xs font-semibold text-slate-900 focus:outline-none focus:border-emerald-500 transition shadow-2xs"
                    />
                  </div>
                </div>
              </div>
            )}

            {/* Tab 2: Exam Session */}
            {activeTab === "session" && (
              <div className="space-y-4">
                {/* Switch existing sessions */}
                <div className="p-3.5 rounded-2xl bg-white border border-[#E8E2D4] space-y-2">
                  <span className="text-xs font-bold text-slate-700 block">
                    Choose Pre-Configured Session Slot:
                  </span>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                    {sessions.map((sess) => (
                      <button
                        key={sess.id}
                        type="button"
                        onClick={() => switchSession(sess.id)}
                        className={`p-2.5 rounded-xl text-left text-xs font-semibold border transition flex items-center justify-between ${
                          sess.id === activeSession.id
                            ? "bg-emerald-50 border-emerald-300 text-emerald-950 font-bold"
                            : "bg-[#FAF8F3] border-[#E8E2D4] text-slate-700 hover:bg-[#F3EFE6]"
                        }`}
                      >
                        <div>
                          <div className="truncate font-bold">{sess.title}</div>
                          <div className="text-[10px] text-slate-500 font-mono mt-0.5">
                            {sess.date} • {sess.timing}
                          </div>
                        </div>
                        {sess.id === activeSession.id && (
                          <span className="h-5 w-5 rounded-full bg-emerald-600 text-white flex items-center justify-center flex-shrink-0">
                            <Check className="h-3 w-3" />
                          </span>
                        )}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Edit active session */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-2">
                  <div className="sm:col-span-2 space-y-1.5">
                    <label className="text-xs font-bold text-slate-700">
                      Exam Session Title
                    </label>
                    <input
                      type="text"
                      value={sessionTitle}
                      onChange={(e) => setSessionTitle(e.target.value)}
                      placeholder="e.g. End-Semester Major Theory Examination • Slot 1"
                      className="w-full px-3.5 py-2.5 rounded-xl bg-white border border-[#E0D9CB] text-xs font-semibold text-slate-900 focus:outline-none focus:border-emerald-500 transition shadow-2xs"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-700">
                      Day / Date Representation
                    </label>
                    <input
                      type="text"
                      value={sessionDate}
                      onChange={(e) => setSessionDate(e.target.value)}
                      placeholder="e.g. Day 1 (Morning Session) or 15-Nov-2026"
                      className="w-full px-3.5 py-2.5 rounded-xl bg-white border border-[#E0D9CB] text-xs font-semibold text-slate-900 focus:outline-none focus:border-emerald-500 transition shadow-2xs"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-bold text-slate-700">
                      Session Shift / Slot
                    </label>
                    <input
                      type="text"
                      value={sessionSlot}
                      onChange={(e) => setSessionSlot(e.target.value)}
                      placeholder="e.g. Slot 1 (Morning) or Slot 2 (Afternoon)"
                      className="w-full px-3.5 py-2.5 rounded-xl bg-white border border-[#E0D9CB] text-xs font-semibold text-slate-900 focus:outline-none focus:border-emerald-500 transition shadow-2xs"
                    />
                  </div>

                  <div className="sm:col-span-2 space-y-1.5">
                    <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                      <Clock className="h-3.5 w-3.5 text-emerald-600" />
                      <span>Exam Timings & Duration</span>
                    </label>
                    <input
                      type="text"
                      value={sessionTiming}
                      onChange={(e) => setSessionTiming(e.target.value)}
                      placeholder="e.g. 09:30 AM - 12:30 PM (3.0 Hours)"
                      className="w-full px-3.5 py-2.5 rounded-xl bg-white border border-[#E0D9CB] text-xs font-semibold text-slate-900 focus:outline-none focus:border-emerald-500 transition shadow-2xs"
                    />
                  </div>
                </div>
              </div>
            )}

            {/* Tab 3: Presets */}
            {activeTab === "presets" && (
              <div className="space-y-3">
                <p className="text-xs text-slate-500">
                  Select a template to instantly apply generalized institute names, codes, exam titles, and dates.
                </p>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  {presetInstitutions.map((item, idx) => (
                    <div
                      key={idx}
                      onClick={() => applyPreset(item)}
                      className="p-4 rounded-2xl bg-white hover:bg-emerald-50/40 border border-[#E8E2D4] hover:border-emerald-300 transition cursor-pointer space-y-2 group shadow-xs"
                    >
                      <div className="flex items-center justify-between">
                        <span className="px-2 py-0.5 rounded-md text-[10px] font-mono font-bold bg-emerald-100 text-emerald-800 border border-emerald-300">
                          {item.badge}
                        </span>
                        <span className="text-xs text-emerald-600 font-bold group-hover:underline">
                          Apply Preset →
                        </span>
                      </div>

                      <div className="font-extrabold text-xs text-slate-900">
                        {item.name}
                      </div>

                      <div className="text-[11px] text-slate-600 space-y-0.5">
                        <div>Exam: {item.examTitle}</div>
                        <div className="text-[10px] text-slate-400 font-mono">
                          {item.code} • {item.date} • {item.timing}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Live Preview Card */}
            <div className="p-4 rounded-2xl bg-white border border-[#E0D9CB] shadow-xs space-y-2">
              <div className="text-[10px] font-mono uppercase tracking-widest text-emerald-800 font-bold flex items-center gap-1.5">
                <FileCheck2 className="h-3.5 w-3.5 text-emerald-600" />
                <span>Live Document & Banner Preview</span>
              </div>

              <div className="p-3 rounded-xl bg-[#FAF8F3] border border-[#E8E2D4] text-center space-y-1">
                <div className="text-xs font-extrabold text-slate-900 uppercase tracking-wide">
                  {name.toUpperCase() || "INSTITUTE NAME"}
                </div>
                <div className="text-[11px] font-bold text-emerald-800 uppercase">
                  {sessionTitle.toUpperCase() || "EXAMINATION TITLE"} • {academicYear.toUpperCase()}
                </div>
                <div className="text-[10px] text-slate-600 font-mono">
                  CODE: {code} • DATE: {sessionDate} • TIMING: {sessionTiming}
                </div>
              </div>
            </div>
          </div>

          {/* Modal Footer Actions */}
          <div className="p-4 sm:p-5 bg-white border-t border-slate-200 flex items-center justify-between gap-3">
            <button
              type="button"
              onClick={() => {
                setName("Apex University of Technology & Management");
                setCode("AUTM-2026");
                setCenterCode("EXAM-CTR-01");
                setAcademicYear("Academic Session 2025-2026");
                setSemester("Semester Examination / Term Assessment");
                setSuperintendent("Dr. Controller of Examinations");
                setSessionTitle("End-Semester Major Theory Examination • Slot 1");
                setSessionDate("Day 1 (Morning Session)");
                setSessionTiming("09:30 AM - 12:30 PM (3.0 Hours)");
                setSessionSlot("Slot 1 (Morning)");
              }}
              className="px-3.5 py-2 rounded-full text-xs font-semibold text-slate-500 hover:text-slate-900 hover:bg-slate-100 transition flex items-center gap-1.5"
            >
              <RotateCcw className="h-3.5 w-3.5" />
              <span>Reset Fields</span>
            </button>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => setIsConfigModalOpen(false)}
                className="px-4 py-2 rounded-full border border-slate-200 bg-white hover:bg-slate-100 text-slate-700 text-xs font-bold transition shadow-2xs"
              >
                Cancel
              </button>

              <button
                type="button"
                onClick={() => handleSave()}
                className="px-6 py-2.5 rounded-full bg-[#161618] hover:bg-black text-white text-xs font-bold transition flex items-center gap-1.5 shadow-md"
              >
                <Check className="h-4 w-4 text-[#D4F754]" />
                <span>Apply & Save Profile</span>
              </button>
            </div>
          </div>
        </motion.div>
      </div>
    </AnimatePresence>
  );
};
