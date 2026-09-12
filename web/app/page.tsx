"use client";
import React from "react";
import Link from "next/link";
import {
  Grid3X3,
  Layers,
  Search,
  Printer,
  ShieldCheck,
  Zap,
  ArrowRight,
  Sliders,
  LayoutGrid,
  Building2,
  QrCode,
  UserCheck,
  Box,
  CheckCircle2,
  Calendar,
} from "lucide-react";
import { useSeating } from "@/lib/context/SeatingContext";
import { DropZone } from "@/components/upload/DropZone";

export default function HomePage() {
  const { metrics, activeSession, collegeProfile, setIsConfigModalOpen } = useSeating();

  const featureCards = [
    {
      title: "Interactive 2D Seating Studio",
      description:
        "Visualize classroom benches in real time. Drag, swap, and inspect student seat assignments with zero-delay updates.",
      link: "/studio",
      icon: <Grid3X3 className="h-6 w-6 text-slate-900" />,
      badge: "Core Studio",
      accent: "bg-[#D4F754]",
    },
    {
      title: "Live QR & Photo Attendance Scanner",
      description:
        "Scan desk QR codes with mobile/webcam to verify student university photos and record timestamps instantly.",
      link: "/scanner",
      icon: <QrCode className="h-6 w-6 text-slate-900" />,
      badge: "Biometric ID",
      accent: "bg-[#B8B5FF]",
    },
    {
      title: "2D Visual Room Blueprint Builder",
      description:
        "Design lecture halls, auditoriums, and computer labs with customizable bench sizes (1-4 seats), doors, and podiums.",
      link: "/builder",
      icon: <Layers className="h-6 w-6 text-slate-900" />,
      badge: "Drag & Drop",
      accent: "bg-[#D4F754]",
    },
    {
      title: "Print Station & Section-Wise Exporter",
      description:
        "1-click batch export for section-wise signature books (F-1/S-1/T-1), A4 door notice charts, and multi-sheet Excel files.",
      link: "/export",
      icon: <Printer className="h-6 w-6 text-slate-900" />,
      badge: "Export",
      accent: "bg-[#D4F754]",
    },
    {
      title: "Faculty Proctor Roster & Incident Log",
      description:
        "Anti-bias invigilator assignment, shift rotations, and live logging of extra answer sheets or malpractice notes.",
      link: "/invigilators",
      icon: <UserCheck className="h-6 w-6 text-slate-900" />,
      badge: "Integrity",
      accent: "bg-[#B8B5FF]",
    },
  ];

  return (
    <div className="space-y-8">
      {/* Hero Bento Header Banner */}
      <section className="bento-card p-8 sm:p-12 relative overflow-hidden bg-gradient-to-br from-white via-white to-[#F2F3F5]">
        <div className="max-w-3xl space-y-4">
          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-[#D4F754] text-black text-xs font-black tracking-wide">
            <LayoutGrid className="h-3.5 w-3.5 fill-black stroke-black" />
            <span>DeskMatrix Seating Architecture</span>
          </div>

          <h1 className="text-4xl sm:text-5xl lg:text-6xl font-black tracking-tight text-slate-900 leading-tight">
            Exam Seating without Conflicts. Automated in Seconds.
          </h1>

          <p className="text-sm sm:text-base text-slate-600 leading-relaxed max-w-2xl font-medium">
            Designed for real-world Mid-Semester Tests (MST) and End-Semester university exams.
            Features multi-branch interleaving, live QR attendance scanners, and section-wise publishing.
          </p>

          <div className="pt-4 flex items-center gap-3 flex-wrap">
            <Link
              href="/studio"
              className="px-6 py-3 rounded-full bg-[#161618] hover:bg-black text-white text-xs font-extrabold tracking-wide transition shadow-md flex items-center gap-2"
            >
              <span>Launch Seating Studio</span>
              <ArrowRight className="h-4 w-4 text-[#D4F754]" />
            </Link>

            <Link
              href="/dashboard"
              className="px-6 py-3 rounded-full bg-white hover:bg-slate-100 text-slate-900 border border-slate-200 text-xs font-bold transition shadow-2xs"
            >
              Open Dashboard
            </Link>

            <button
              onClick={() => setIsConfigModalOpen(true)}
              className="px-4 py-3 rounded-full bg-[#F2F3F5] hover:bg-slate-200 text-slate-700 text-xs font-bold transition"
            >
              ⚙️ Configure College & Exam
            </button>
          </div>
        </div>
      </section>

      {/* Quick Metrics Bar */}
      <section className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <div className="bento-card p-5 text-center">
          <div className="text-3xl font-black text-slate-900 font-mono">
            {metrics.totalRooms}
          </div>
          <div className="text-xs text-slate-400 font-bold uppercase tracking-wider mt-1">
            Exam Halls Ready
          </div>
        </div>

        <div className="bento-card p-5 text-center">
          <div className="text-3xl font-black text-slate-900 font-mono">
            {metrics.seatedStudents}
          </div>
          <div className="text-xs text-slate-400 font-bold uppercase tracking-wider mt-1">
            Candidates Seated
          </div>
        </div>

        <div className="bento-card p-5 text-center">
          <div className="text-3xl font-black text-slate-900 font-mono flex items-center justify-center gap-1">
            <span>{metrics.conflictFreeRate}%</span>
            <span className="h-2 w-2 rounded-full bg-[#D4F754]" />
          </div>
          <div className="text-xs text-slate-400 font-bold uppercase tracking-wider mt-1">
            Zero Conflict Score
          </div>
        </div>

        <div className="bento-card p-5 text-center">
          <div className="text-3xl font-black text-slate-900 font-mono">
            {metrics.utilizationRate}%
          </div>
          <div className="text-xs text-slate-400 font-bold uppercase tracking-wider mt-1">
            Hall Utilization
          </div>
        </div>
      </section>

      {/* Excel Drag & Drop Upload Bento Card */}
      <section className="bento-card p-6 sm:p-8 space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-4">
          <div>
            <h2 className="text-lg font-black text-slate-900">
              Upload Institutional Layouts & Roll Lists
            </h2>
            <p className="text-xs text-slate-500">
              Drop your custom Excel spreadsheets to instantly re-calculate optimal seating plans.
            </p>
          </div>
        </div>
        <DropZone />
      </section>

      {/* Feature Grid Bento Cards */}
      <section className="space-y-4">
        <div className="flex items-center justify-between">
          <h2 className="text-xl font-black text-slate-900">
            Examination Modules
          </h2>
          <span className="text-xs text-slate-400 font-mono">
            6 Specialized Workspaces
          </span>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {featureCards.map((card, idx) => (
            <Link
              key={idx}
              href={card.link}
              className="bento-card p-5 flex flex-col justify-between space-y-4 group hover:border-slate-300"
            >
              <div className="flex items-center justify-between">
                <div className="h-10 w-10 rounded-2xl bg-slate-100 group-hover:bg-[#D4F754] transition flex items-center justify-center">
                  {card.icon}
                </div>
                <span className="px-2.5 py-0.5 rounded-full bg-slate-100 text-slate-700 text-[10px] font-bold font-mono">
                  {card.badge}
                </span>
              </div>

              <div>
                <h3 className="font-extrabold text-sm text-slate-900 group-hover:text-black">
                  {card.title}
                </h3>
                <p className="text-xs text-slate-500 mt-1 line-clamp-2">
                  {card.description}
                </p>
              </div>

              <div className="pt-2 border-t border-slate-100 flex items-center justify-between text-xs font-bold text-slate-900 group-hover:underline">
                <span>Launch Workspace</span>
                <span>→</span>
              </div>
            </Link>
          ))}
        </div>
      </section>
    </div>
  );
}
