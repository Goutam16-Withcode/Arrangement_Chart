"use client";
import React from "react";
import Link from "next/link";
import { Spotlight } from "@/components/ui/spotlight";
import { SparklesCore } from "@/components/ui/sparkles";
import { Button as MovingBorderButton } from "@/components/ui/moving-border";
import { HoverEffect } from "@/components/ui/card-hover-effect";
import { DropZone } from "@/components/upload/DropZone";
import {
  Grid3X3,
  Layers,
  Search,
  Printer,
  ShieldCheck,
  Zap,
  ArrowRight,
  Sliders,
  Sparkles,
  QrCode,
  UserCheck,
  Box,
} from "lucide-react";
import { useSeating } from "@/lib/context/SeatingContext";

export default function HomePage() {
  const { metrics, activeSession, collegeProfile } = useSeating();

  const featureCards = [
    {
      title: "Interactive 2D Seating Studio",
      description:
        "Visualize classroom benches in real time. Drag, swap, and inspect student seat assignments with zero-delay updates.",
      link: "/studio",
      icon: <Grid3X3 className="h-6 w-6 text-emerald-600" />,
      badge: "Core Studio",
    },
    {
      title: "Live QR & Photo Attendance Scanner",
      description:
        "Scan desk QR codes with mobile/webcam to verify student university photos and record timestamps instantly.",
      link: "/scanner",
      icon: <QrCode className="h-6 w-6 text-emerald-600" />,
      badge: "Biometric ID",
    },
    {
      title: "2D Visual Room Blueprint Builder",
      description:
        "Design lecture halls, auditoriums, and computer labs with customizable bench sizes (1-4 seats), doors, and podiums.",
      link: "/builder",
      icon: <Layers className="h-6 w-6 text-emerald-600" />,
      badge: "Drag & Drop",
    },
    {
      title: "Student Kiosk & Indoor Wayfinding",
      description:
        "Instant roll number search portal with 3D pin classroom locators, digital admit pass, and turn-by-turn indoor routing.",
      link: "/find-seat",
      icon: <Search className="h-6 w-6 text-emerald-600" />,
      badge: "Self Service",
    },
    {
      title: "Print Station & Section-Wise Exporter",
      description:
        "1-click batch export for section-wise signature books (F-1/S-1/T-1), A4 door notice charts, and multi-sheet Excel files.",
      link: "/export",
      icon: <Printer className="h-6 w-6 text-emerald-600" />,
      badge: "Export",
    },
    {
      title: "Faculty Proctor Roster & Incident Log",
      description:
        "Anti-bias invigilator assignment, shift rotations, and live logging of extra answer sheets or malpractice notes.",
      link: "/invigilators",
      icon: <UserCheck className="h-6 w-6 text-emerald-600" />,
      badge: "Integrity",
    },
  ];

  return (
    <div className="relative overflow-hidden min-h-screen bg-[#FBF9F4]">
      {/* Background Spotlight with Light Green Mint glow */}
      <Spotlight
        className="-top-40 left-0 md:left-60 md:-top-20"
        fill="#34d399"
      />

      {/* Hero Section */}
      <section className="relative pt-16 pb-14 px-4 sm:px-6 lg:px-8 max-w-7xl mx-auto z-10 text-center">
        <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-semibold mb-8 animate-pulse-glow shadow-xs">
          <Sparkles className="h-3.5 w-3.5 text-emerald-600" />
          <span>⚡ Universal Exam Seating & Attendance Automation Suite</span>
        </div>

        <h1 className="text-4xl sm:text-6xl lg:text-7xl font-extrabold tracking-tight text-slate-900 max-w-4xl mx-auto leading-tight">
          Exam Seating without{" "}
          <span className="text-gradient">Conflicts</span>. Automated in{" "}
          <span className="text-gradient-cyan">Seconds</span>.
        </h1>

        <p className="mt-6 text-base sm:text-lg text-slate-600 max-w-2xl mx-auto leading-relaxed font-normal">
          Designed for real-world Mid-Semester Tests (MST) and End-Semester university exams.
          Features multi-branch interleaving, live QR attendance scanners, and section-wise publishing.
        </p>

        {/* CTA Buttons */}
        <div className="mt-10 flex flex-wrap items-center justify-center gap-4">
          <Link href="/studio">
            <MovingBorderButton
              borderRadius="1.25rem"
              className="bg-emerald-600 text-white border-emerald-500 px-6 py-3 font-bold text-sm flex items-center gap-2 hover:bg-emerald-700 shadow-md shadow-emerald-600/25 transition"
            >
              <span>Launch Seating Studio</span>
              <ArrowRight className="h-4 w-4 text-emerald-100" />
            </MovingBorderButton>
          </Link>

          <Link
            href="/dashboard"
            className="px-6 py-3 rounded-2xl bg-white hover:bg-emerald-50/50 text-slate-800 border border-[#E0D9CB] text-sm font-semibold transition flex items-center gap-2 shadow-xs"
          >
            <Sliders className="h-4 w-4 text-emerald-700" />
            <span>Open Dashboard</span>
          </Link>
        </div>

        {/* Sparkles Particle Divider */}
        <div className="w-full h-20 relative mt-12">
          <div className="absolute inset-x-20 top-0 bg-gradient-to-r from-transparent via-emerald-400 to-transparent h-[2px] w-3/4 blur-xs" />
          <div className="absolute inset-x-20 top-0 bg-gradient-to-r from-transparent via-emerald-500 to-transparent h-px w-3/4" />

          <SparklesCore
            background="transparent"
            minSize={0.4}
            maxSize={1.8}
            particleDensity={60}
            className="w-full h-full"
            particleColor="#059669"
          />

          <div className="absolute inset-0 w-full h-full bg-[#FBF9F4] [mask-image:radial-gradient(350px_200px_at_top,transparent_20%,white)]"></div>
        </div>
      </section>

      {/* Live Metrics Showcase */}
      <section className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mb-16">
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div className="p-5 rounded-2xl bg-white border border-[#E8E2D4] shadow-xs text-center">
            <div className="text-3xl font-extrabold text-slate-900 font-mono">
              {metrics.totalRooms}
            </div>
            <div className="text-xs text-slate-500 mt-1 uppercase tracking-wider font-semibold">
              Exam Halls Ready
            </div>
          </div>

          <div className="p-5 rounded-2xl bg-white border border-[#E8E2D4] shadow-xs text-center">
            <div className="text-3xl font-extrabold text-emerald-600 font-mono">
              {metrics.seatedStudents}
            </div>
            <div className="text-xs text-slate-500 mt-1 uppercase tracking-wider font-semibold">
              Candidates Seated
            </div>
          </div>

          <div className="p-5 rounded-2xl bg-white border border-[#E8E2D4] shadow-xs text-center">
            <div className="text-3xl font-extrabold text-teal-600 font-mono">
              {metrics.conflictFreeRate}%
            </div>
            <div className="text-xs text-slate-500 mt-1 uppercase tracking-wider font-semibold">
              Zero Conflict Score
            </div>
          </div>

          <div className="p-5 rounded-2xl bg-white border border-[#E8E2D4] shadow-xs text-center">
            <div className="text-3xl font-extrabold text-lime-700 font-mono">
              {metrics.utilizationRate}%
            </div>
            <div className="text-xs text-slate-500 mt-1 uppercase tracking-wider font-semibold">
              Hall Utilization
            </div>
          </div>
        </div>
      </section>

      {/* Quick Excel Ingestion Dropzone */}
      <section className="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8 mb-20">
        <div className="text-center mb-6">
          <h2 className="text-2xl font-bold text-slate-900 tracking-tight">
            Upload Layout & Roll Lists
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            Drop your institutional spreadsheets to instantly calculate optimal seat plans.
          </p>
        </div>
        <DropZone />
      </section>

      {/* Feature Navigation Grid (CardHoverEffect) */}
      <section className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mb-24">
        <div className="text-center mb-8">
          <h2 className="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">
            Comprehensive Examination Modules
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            Everything you need for seamless MST and End-Sem management in one unified platform.
          </p>
        </div>

        <HoverEffect items={featureCards} />
      </section>
    </div>
  );
}
