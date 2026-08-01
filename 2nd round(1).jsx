import React, { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, RadarChart,
  PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar, Brush
} from "recharts";
import {
  MapPin, AlertTriangle, CheckCircle2, Clock, Users, Camera, Search,
  BarChart3, Map as MapIcon, Bell, Globe, Moon, Sun, LogIn, UserPlus,
  Shield, Truck, ClipboardList, TrendingUp, Download, FileSpreadsheet,
  FileText, X, ChevronRight, Menu, Home, LayoutDashboard, Brain,
  Recycle, Building2, Trash2, MessageCircle, Send, ThermometerSun,
  Droplets, Store, CalendarDays, Filter, ArrowUpRight, ArrowDownRight, ArrowLeft,
  Info, Settings, LogOut, Star, Navigation, Layers, ChevronDown, Plus,
  Mail, Lock, Phone, Eye, EyeOff, Copy, Check
} from "lucide-react";

/* ---------------------------------- Demo / Seed Accounts ---------------------------------- */
// Citizens sign in with EMAIL. Officers sign in with a unique Officer UID, and
// live in a separate "officers database" below.
const DEMO_ACCOUNTS = [
  { name: "A. Sravani", email: "citizen@gvmc.gov.in", phone: "9440011111", password: "citizen123", role: "citizen" },
];
const DEMO_OFFICERS = [
  { uid: "GVMC-OFF-001", name: "Ravi Prasad", phone: "9440099999", password: "admin123", role: "admin", designation: "Sanitation Officer" },
];

// Generates the next sequential Officer UID, e.g. GVMC-OFF-002.
function nextOfficerUID(officers) {
  const nums = officers.map(o => parseInt((String(o.uid).match(/(\d+)$/) || [])[1] || "0", 10));
  const next = (nums.length ? Math.max(...nums) : 0) + 1;
  return `GVMC-OFF-${String(next).padStart(3, "0")}`;
}

/* ---------------------------------- i18n ---------------------------------- */
const STR = {
  en: {
    appName: "Swachha GVMC", tagline: "AI Smart Garbage Vulnerable Point System",
    home: "Home", citizen: "Citizen Portal", admin: "Admin Dashboard", ai: "AI Predictions",
    login: "Login", register: "Register", logout: "Logout",
    heroTitle: "Predict. Prevent. Keep GVMC Clean.",
    heroSub: "An AI + GIS powered platform that helps citizens report garbage vulnerable points and empowers GVMC to predict and prevent recurring dumping hotspots before they happen.",
    reportNow: "Report Garbage", viewMap: "View Live Map",
    totalComplaints: "Total Complaints", activeGVPs: "Active GVPs", cleanedSpots: "Cleaned Spots", highRisk: "High-Risk Areas",
    overview: "Project Overview",
    overviewText: "GVMC's Smart GVP system combines citizen reporting, machine learning and geospatial analytics to identify Garbage Vulnerable Points (GVPs) across the city, predict where garbage is likely to recur, and route sanitation resources proactively — before spots become chronic dump sites.",
  },
  te: {
    appName: "స్వచ్ఛ GVMC", tagline: "AI స్మార్ట్ చెత్త ప్రమాదకర ప్రాంతాల వ్యవస్థ",
    home: "హోమ్", citizen: "పౌర పోర్టల్", admin: "అడ్మిన్ డాష్‌బోర్డ్", ai: "AI అంచనాలు",
    login: "లాగిన్", register: "నమోదు", logout: "లాగ్ అవుట్",
    heroTitle: "అంచనా వేయండి. నివారించండి. GVMCని పరిశుభ్రంగా ఉంచండి.",
    heroSub: "పౌరులు చెత్త ప్రమాదకర ప్రాంతాలను నివేదించడానికి మరియు GVMC పునరావృత హాట్‌స్పాట్‌లను ముందుగానే అంచనా వేయడానికి సహాయపడే AI + GIS ఆధారిత వేదిక.",
    reportNow: "చెత్తను నివేదించండి", viewMap: "లైవ్ మ్యాప్ చూడండి",
    totalComplaints: "మొత్తం ఫిర్యాదులు", activeGVPs: "సక్రియ GVPలు", cleanedSpots: "శుభ్రం చేసిన ప్రాంతాలు", highRisk: "అధిక-ప్రమాద ప్రాంతాలు",
    overview: "ప్రాజెక్ట్ అవలోకనం",
    overviewText: "GVMC యొక్క స్మార్ట్ GVP వ్యవస్థ పౌర నివేదికలు, మెషిన్ లెర్నింగ్ మరియు జియోస్పేషియల్ విశ్లేషణలను మిళితం చేసి నగరవ్యాప్తంగా చెత్త ప్రమాదకర ప్రాంతాలను గుర్తిస్తుంది.",
  },
};

/* ---------------------------------- Mock Data ---------------------------------- */
const WARDS = ["Ward 12 - MVP Colony", "Ward 15 - Dwaraka Nagar", "Ward 22 - Gajuwaka", "Ward 8 - Seethammadhara", "Ward 31 - Pendurthi", "Ward 19 - Madhurawada", "Ward 5 - Old Town", "Ward 27 - Akkayyapalem"];
const GARBAGE_TYPES = ["Household Waste", "Construction Debris", "Market/Vegetable Waste", "Plastic & Packaging", "E-Waste", "Medical Waste", "Mixed/Dumped Waste"];
const STATUSES = ["Submitted", "Under Verification", "Cleaning Scheduled", "Cleaning in Progress", "Cleaned", "Closed"];
const STATUS_COLORS = {
  "Submitted": { bg: "#EAF3DE", text: "#3B6D11", dot: "#639922" },
  "Under Verification": { bg: "#FAEEDA", text: "#854F0B", dot: "#EF9F27" },
  "Cleaning Scheduled": { bg: "#E6F1FB", text: "#0C447C", dot: "#378ADD" },
  "Cleaning in Progress": { bg: "#EDE9FE", text: "#4C1D95", dot: "#7C3AED" },
  "Cleaned": { bg: "#E1F5EE", text: "#085041", dot: "#1D9E75" },
  "Closed": { bg: "#F1EFE8", text: "#444441", dot: "#888780" },
};
const WORKERS = [
  { id: "W-101", name: "Ravi Kumar", ward: "Ward 12", phone: "9440012345", status: "On Duty", tasksToday: 4 },
  { id: "W-102", name: "Lakshmi Prasad", ward: "Ward 15", phone: "9440012346", status: "On Duty", tasksToday: 3 },
  { id: "W-103", name: "Suresh Babu", ward: "Ward 22", phone: "9440012347", status: "Off Duty", tasksToday: 0 },
  { id: "W-104", name: "Anitha Rao", ward: "Ward 8", phone: "9440012348", status: "On Duty", tasksToday: 5 },
  { id: "W-105", name: "Mahesh Naidu", ward: "Ward 31", phone: "9440012349", status: "On Duty", tasksToday: 2 },
];

function seededRandom(seed) {
  let s = seed;
  return () => { s = (s * 9301 + 49297) % 233280; return s / 233280; };
}
const rnd = seededRandom(42);

const HOTSPOTS = [
  { id: "GVP-001", name: "Beach Road Junction Market", ward: WARDS[0], lat: 17.7231, lng: 83.3244, risk: "High", complaints: 34, lastCleaned: "2026-07-24", recurrence: 87, rootCause: "Weekly market waste + irregular collection", nearby: ["3 dustbins (0.9km avg)", "Fish market 120m", "2 bus stops"] },
  { id: "GVP-002", name: "Dwaraka Nagar Bus Stand", ward: WARDS[1], lat: 17.7315, lng: 83.3037, risk: "High", complaints: 29, lastCleaned: "2026-07-22", recurrence: 81, rootCause: "High footfall, no covered bins", nearby: ["Bus terminal", "5 shops", "1 dustbin (1.4km)"] },
  { id: "GVP-003", name: "Gajuwaka Signal Vacant Plot", ward: WARDS[2], lat: 17.6868, lng: 83.2185, risk: "High", complaints: 41, lastCleaned: "2026-07-18", recurrence: 92, rootCause: "Illegal dumping on vacant land, low surveillance", nearby: ["Vacant plot", "Industrial area 200m"] },
  { id: "GVP-004", name: "Seethammadhara Park Corner", ward: WARDS[3], lat: 17.7412, lng: 83.3128, risk: "Medium", complaints: 14, lastCleaned: "2026-07-27", recurrence: 54, rootCause: "Residential mixed waste, missed pickup 1x/week", nearby: ["Park", "2 dustbins nearby"] },
  { id: "GVP-005", name: "Pendurthi Main Road", ward: WARDS[4], lat: 17.8038, lng: 83.2354, risk: "Medium", complaints: 11, lastCleaned: "2026-07-26", recurrence: 48, rootCause: "Roadside vendor waste accumulation", nearby: ["Vendor stalls", "1 dustbin 1.1km"] },
  { id: "GVP-006", name: "Madhurawada IT Hub Lane", ward: WARDS[5], lat: 17.8072, lng: 83.3782, risk: "Low", complaints: 5, lastCleaned: "2026-07-29", recurrence: 21, rootCause: "Occasional construction debris", nearby: ["IT park", "Covered bin 300m"] },
  { id: "GVP-007", name: "Old Town Temple Street", ward: WARDS[6], lat: 17.7015, lng: 83.2945, risk: "Medium", complaints: 17, lastCleaned: "2026-07-25", recurrence: 58, rootCause: "Flower/temple waste + narrow street access", nearby: ["Temple", "Market 150m"] },
  { id: "GVP-008", name: "Akkayyapalem Circle", ward: WARDS[7], lat: 17.7398, lng: 83.2989, risk: "Low", complaints: 6, lastCleaned: "2026-07-29", recurrence: 19, rootCause: "Low density, well-serviced route", nearby: ["2 dustbins", "Residential"] },
];

function genComplaints() {
  const list = [];
  for (let i = 1; i <= 42; i++) {
    const w = WARDS[Math.floor(rnd() * WARDS.length)];
    const status = STATUSES[Math.floor(rnd() * STATUSES.length)];
    const type = GARBAGE_TYPES[Math.floor(rnd() * GARBAGE_TYPES.length)];
    const day = 1 + Math.floor(rnd() * 29);
    list.push({
      id: `GVMC-${2026000 + i}`,
      ward: w,
      type,
      status,
      address: `${5 + i} Cross Road, ${w.split(" - ")[1]}`,
      date: `2026-07-${String(day).padStart(2, "0")}`,
      citizen: ["A. Sravani", "M. Karthik", "P. Bhavana", "K. Ravi Teja", "S. Divya"][i % 5],
      description: "Garbage accumulated near roadside, not collected for several days.",
      assignedTo: rnd() > 0.4 ? WORKERS[Math.floor(rnd() * WORKERS.length)].name : null,
    });
  }
  return list;
}
const COMPLAINTS_SEED = genComplaints();

const TREND_DATA = [
  { month: "Feb", complaints: 210, cleaned: 180 },
  { month: "Mar", complaints: 245, cleaned: 220 },
  { month: "Apr", complaints: 198, cleaned: 190 },
  { month: "May", complaints: 267, cleaned: 230 },
  { month: "Jun", complaints: 289, cleaned: 265 },
  { month: "Jul", complaints: 253, cleaned: 241 },
];

const WARD_ANALYSIS = WARDS.map((w, i) => ({
  ward: w.split(" - ")[1],
  complaints: [34, 29, 41, 14, 11, 5, 17, 6][i],
  risk: [88, 81, 92, 54, 48, 21, 58, 19][i],
}));

const ROOT_CAUSE_DATA = [
  { name: "Irregular Collection", value: 34 },
  { name: "Market Waste", value: 26 },
  { name: "Illegal Dumping", value: 21 },
  { name: "Insufficient Bins", value: 12 },
  { name: "Construction Debris", value: 7 },
];
const PIE_COLORS = ["#0A4C8C", "#0B5D3B", "#D6432B", "#E8A93B", "#7F77DD"];

const FEATURE_IMPORTANCE = [
  { feature: "Past Complaint Freq.", importance: 92 },
  { feature: "Distance to Market", importance: 78 },
  { feature: "Collection Frequency", importance: 74 },
  { feature: "Population Density", importance: 65 },
  { feature: "Dustbin Density", importance: 61 },
  { feature: "Rainfall (monsoon)", importance: 44 },
  { feature: "Days Since Cleaned", importance: 83 },
];

/* ---------------------------------- Small UI atoms ---------------------------------- */
function RiskBadge({ risk }) {
  const map = {
    High: { bg: "#FCEBEB", text: "#791F1F", dot: "#E24B4A" },
    Medium: { bg: "#FAEEDA", text: "#854F0B", dot: "#EF9F27" },
    Low: { bg: "#EAF3DE", text: "#27500A", dot: "#639922" },
  };
  const c = map[risk] || map.Low;
  return (
    <span style={{ background: c.bg, color: c.text }} className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold">
      <span style={{ background: c.dot }} className="w-1.5 h-1.5 rounded-full"></span>
      {risk} Risk
    </span>
  );
}

function StatusPill({ status }) {
  const c = STATUS_COLORS[status] || STATUS_COLORS.Submitted;
  return (
    <span style={{ background: c.bg, color: c.text }} className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold whitespace-nowrap">
      <span style={{ background: c.dot }} className="w-1.5 h-1.5 rounded-full"></span>
      {status}
    </span>
  );
}

/* ---------------------------------- Live Complaint Tracker ---------------------------------- */
// Horizontal stepper showing a complaint's journey from submission to completion.
// It reads the complaint's current status (shared store), so when an officer changes
// the status in the Officer Portal, this re-renders and advances automatically.
const TRACK_STAGES = ["Submitted", "Under Verification", "Cleaning Scheduled", "Cleaning in Progress", "Cleaned", "Closed"];
const TRACK_SHORT = { "Submitted": "Filed", "Under Verification": "Verifying", "Cleaning Scheduled": "Scheduled", "Cleaning in Progress": "In Progress", "Cleaned": "Cleaned", "Closed": "Closed" };

function ComplaintTracker({ complaint, dark }) {
  const currentIdx = Math.max(0, TRACK_STAGES.indexOf(complaint.status));
  const histAt = {};
  (complaint.statusHistory || []).forEach(h => { histAt[h.status] = h.at; });
  const fmt = (iso) => { try { return new Date(iso).toLocaleString("en-IN", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }); } catch { return ""; } };
  const doneColor = "#0B5D3B";
  const idleColor = dark ? "#334155" : "#cbd5e1";
  const resolved = complaint.status === "Cleaned" || complaint.status === "Closed";

  return (
    <div className="mt-3">
      <div className="flex items-start">
        {TRACK_STAGES.map((st, i) => {
          const done = i < currentIdx;
          const current = i === currentIdx;
          const active = done || current;
          const at = histAt[st];
          return (
            <React.Fragment key={st}>
              <div className="flex flex-col items-center text-center" style={{ flex: "0 0 auto", width: 62 }}>
                <div className="relative w-6 h-6 rounded-full flex items-center justify-center" style={{ background: active ? doneColor : idleColor }}>
                  {done ? (
                    <Check size={13} color="#fff" />
                  ) : current ? (
                    <>
                      <span className="w-2 h-2 rounded-full bg-white"></span>
                      {!resolved && <span className="absolute inset-0 rounded-full animate-ping" style={{ background: doneColor, opacity: 0.4 }}></span>}
                    </>
                  ) : (
                    <span className="text-[10px] font-bold" style={{ color: dark ? "#94a3b8" : "#64748b" }}>{i + 1}</span>
                  )}
                </div>
                <span className={`text-[9px] mt-1 leading-tight font-semibold ${active ? (dark ? "text-white" : "text-slate-800") : (dark ? "text-slate-500" : "text-slate-400")}`}>{TRACK_SHORT[st]}</span>
                <span className={`text-[8px] leading-tight ${dark ? "text-slate-500" : "text-slate-400"}`}>{at ? fmt(at) : ""}</span>
              </div>
              {i < TRACK_STAGES.length - 1 && (
                <div className="h-0.5 mt-3 rounded-full" style={{ flex: "1 1 auto", background: i < currentIdx ? doneColor : idleColor }} />
              )}
            </React.Fragment>
          );
        })}
      </div>
    </div>
  );
}


function StatCard({ icon: Icon, label, value, sub, accent, dark }) {
  return (
    <div className={`rounded-2xl p-5 border ${dark ? "bg-slate-800/60 border-slate-700" : "bg-white border-slate-200"} shadow-sm hover:shadow-md transition-shadow`}>
      <div className="flex items-start justify-between">
        <div>
          <p className={`text-xs font-medium uppercase tracking-wide ${dark ? "text-slate-400" : "text-slate-500"}`}>{label}</p>
          <p className={`text-3xl font-bold mt-1.5 ${dark ? "text-white" : "text-slate-900"}`}>{value}</p>
          {sub && <p className="text-xs mt-1.5 flex items-center gap-1" style={{ color: accent }}>{sub}</p>}
        </div>
        <div style={{ background: accent + "1a" }} className="p-2.5 rounded-xl">
          <Icon size={20} style={{ color: accent }} />
        </div>
      </div>
    </div>
  );
}

function SectionTitle({ eyebrow, title, sub, dark }) {
  return (
    <div className="mb-6">
      {eyebrow && <p className="text-xs font-bold tracking-widest uppercase mb-1" style={{ color: "#0B5D3B" }}>{eyebrow}</p>}
      <h2 className={`text-2xl font-bold ${dark ? "text-white" : "text-slate-900"}`}>{title}</h2>
      {sub && <p className={`text-sm mt-1.5 max-w-2xl ${dark ? "text-slate-400" : "text-slate-500"}`}>{sub}</p>}
    </div>
  );
}

function Card({ children, className = "", dark }) {
  return (
    <div className={`rounded-2xl border ${dark ? "bg-slate-800/60 border-slate-700" : "bg-white border-slate-200"} shadow-sm ${className}`}>
      {children}
    </div>
  );
}

/* ---------------------------------- India GVP Map (Leaflet + OpenStreetMap) ---------------------------------- */
const RISK_COLOR = { High: "#E24B4A", Medium: "#EF9F27", Low: "#639922" };
const INDIA_CENTER = [20.5937, 78.9629];
const INDIA_ZOOM = 5;

// Loads Leaflet (CSS + JS) from a CDN once, shared across every map instance.
let _leafletPromise = null;
function loadLeaflet() {
  if (typeof window !== "undefined" && window.L) return Promise.resolve(window.L);
  if (_leafletPromise) return _leafletPromise;
  _leafletPromise = new Promise((resolve, reject) => {
    if (!document.getElementById("leaflet-css")) {
      const link = document.createElement("link");
      link.id = "leaflet-css";
      link.rel = "stylesheet";
      link.href = "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css";
      document.head.appendChild(link);
    }
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js";
    script.async = true;
    script.onload = () => resolve(window.L);
    script.onerror = () => reject(new Error("Failed to load Leaflet from CDN"));
    document.body.appendChild(script);
  });
  return _leafletPromise;
}

function HotspotMap({ dark, onSelect, selected, compact }) {
  const containerRef = useRef(null);
  const mapRef = useRef(null);
  const markersRef = useRef({});
  const [status, setStatus] = useState("loading"); // loading | ready | error

  // Initialise the Leaflet map once, with OSM + satellite tile layers and GVP markers.
  useEffect(() => {
    let cancelled = false;
    loadLeaflet().then((L) => {
      if (cancelled || !containerRef.current || mapRef.current) return;
      try {
        const map = L.map(containerRef.current, {
          center: INDIA_CENTER,
          zoom: INDIA_ZOOM,
          zoomControl: true,
          scrollWheelZoom: true,
          worldCopyJump: true,
        });
        mapRef.current = map;

        const street = L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
          maxZoom: 19,
          attribution: '&copy; OpenStreetMap contributors',
        }).addTo(map);
        const satellite = L.tileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}", {
          maxZoom: 19,
          attribution: "Tiles &copy; Esri, Maxar, Earthstar Geographics",
        });
        L.control.layers({ "Street": street, "Satellite": satellite }, null, { position: "topright" }).addTo(map);

        // Home / reset-to-India control.
        const HomeCtrl = L.Control.extend({
          options: { position: "topleft" },
          onAdd: function () {
            const btn = L.DomUtil.create("button", "leaflet-bar");
            btn.type = "button";
            btn.title = "Reset to India view";
            btn.innerHTML = "&#8962;"; // house glyph
            Object.assign(btn.style, {
              width: "34px", height: "34px", cursor: "pointer", background: "#fff",
              border: "none", borderRadius: "4px", fontSize: "18px", lineHeight: "34px",
              textAlign: "center", display: "block",
            });
            L.DomEvent.disableClickPropagation(btn);
            L.DomEvent.on(btn, "click", (e) => { L.DomEvent.stop(e); map.setView(INDIA_CENTER, INDIA_ZOOM); });
            return btn;
          },
        });
        map.addControl(new HomeCtrl());

        // GVP markers (vector circleMarkers — no external icon images needed).
        HOTSPOTS.forEach((h) => {
          const marker = L.circleMarker([h.lat, h.lng], {
            radius: h.risk === "High" ? 10 : h.risk === "Medium" ? 8 : 6,
            color: "#ffffff", weight: 2, fillColor: RISK_COLOR[h.risk], fillOpacity: 0.95,
          }).addTo(map);
          marker.bindPopup(
            `<div style="min-width:190px;font-family:system-ui,sans-serif">
               <div style="font-weight:700;font-size:13px;color:#0f172a">${h.name}</div>
               <div style="color:#64748b;font-size:11px;margin-bottom:6px">${h.ward}</div>
               <div style="font-size:12px;color:#0f172a"><b>Risk:</b> ${h.risk} &nbsp;·&nbsp; <b>Complaints:</b> ${h.complaints}</div>
               <div style="font-size:12px;color:#0f172a"><b>Predicted recurrence:</b> ${h.recurrence}%</div>
               <div style="font-size:11px;color:#475569;margin-top:4px">${h.rootCause}</div>
             </div>`
          );
          marker.on("click", () => onSelect && onSelect(h));
          markersRef.current[h.id] = marker;
        });

        setTimeout(() => { if (mapRef.current) mapRef.current.invalidateSize(); }, 120);
        if (!cancelled) setStatus("ready");
      } catch (err) {
        if (!cancelled) setStatus("error");
      }
    }).catch(() => { if (!cancelled) setStatus("error"); });

    return () => {
      cancelled = true;
      if (mapRef.current) { mapRef.current.remove(); mapRef.current = null; }
      markersRef.current = {};
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Fly to a selected GVP and open its popup; Home/Reset returns to the India view.
  useEffect(() => {
    const map = mapRef.current;
    if (!map || status !== "ready" || !selected) return;
    const m = markersRef.current[selected.id];
    if (m) {
      map.flyTo([selected.lat, selected.lng], Math.max(map.getZoom(), 13), { duration: 0.8 });
      m.openPopup();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected?.id, status]);

  // Keep the map correctly sized on container/viewport changes (responsive).
  useEffect(() => {
    const map = mapRef.current;
    if (!map) return;
    const onResize = () => map.invalidateSize();
    window.addEventListener("resize", onResize);
    const t = setTimeout(onResize, 200);
    return () => { window.removeEventListener("resize", onResize); clearTimeout(t); };
  }, [status]);

  const height = compact ? 340 : 460;

  return (
    <div className="relative w-full rounded-xl overflow-hidden border" style={{ borderColor: dark ? "#334155" : "#e2e8f0", isolation: "isolate" }}>
      <div ref={containerRef} style={{ width: "100%", height }} />

      {status !== "ready" && (
        <div className="absolute inset-0 flex items-center justify-center px-6 text-center" style={{ background: dark ? "#0b1622" : "#dcebf5", zIndex: 500 }}>
          {status === "loading" ? (
            <div className="flex items-center gap-2 text-sm" style={{ color: dark ? "#94a3b8" : "#64748b" }}>
              <span className="w-4 h-4 border-2 border-current border-t-transparent rounded-full animate-spin" /> Loading map…
            </div>
          ) : (
            <div className="text-sm" style={{ color: dark ? "#94a3b8" : "#64748b" }}>
              <MapIcon size={22} className="mx-auto mb-2 opacity-70" />
              Map tiles couldn't be loaded in this preview. They'll appear automatically when the app runs with network access.
            </div>
          )}
        </div>
      )}

      {/* Risk legend (non-interactive overlay) */}
      {status === "ready" && (
        <div className="absolute bottom-3 left-3 flex gap-3 text-xs font-medium px-3 py-2 rounded-lg backdrop-blur" style={{ background: dark ? "rgba(15,23,42,0.85)" : "rgba(255,255,255,0.9)", color: dark ? "#e2e8f0" : "#334155", zIndex: 500, pointerEvents: "none" }}>
          <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-full" style={{ background: "#E24B4A" }}></span>High</span>
          <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-full" style={{ background: "#EF9F27" }}></span>Medium</span>
          <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-full" style={{ background: "#639922" }}></span>Low</span>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------- Schedule Crew (inline form) ---------------------------------- */
// Rendered inline inside the detail panel (normal document flow) rather than as a
// fixed overlay — so it can never be blocked by a stacking context, an ancestor
// transform, or a sandboxed preview frame. This is what makes the action reliably
// clickable on both desktop and mobile.
function ScheduleCrewForm({ hotspot, dark, onCancel, onScheduled }) {
  const today = new Date().toISOString().slice(0, 10);
  const [workerId, setWorkerId] = useState("");
  const [date, setDate] = useState(today);
  const [notes, setNotes] = useState("");
  const [status, setStatus] = useState("idle"); // idle | submitting | error
  const [errorMsg, setErrorMsg] = useState("");
  const submitting = status === "submitting";

  const labelCls = `text-xs font-semibold ${dark ? "text-slate-400" : "text-slate-500"}`;
  const inputCls = `w-full mt-1 px-3 py-2.5 rounded-xl border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white" : "bg-white border-slate-200 text-slate-900"}`;
  const lockStyle = submitting ? { opacity: 0.6, pointerEvents: "none" } : undefined;

  const submit = async () => {
    setErrorMsg("");
    if (!workerId) { setStatus("error"); setErrorMsg("Select a sanitation worker to lead the crew."); return; }
    if (!date) { setStatus("error"); setErrorMsg("Choose a cleaning date."); return; }
    setStatus("submitting");
    try {
      // Simulated scheduling request. In production: POST /api/gvp/:id/schedule.
      await new Promise((resolve) => setTimeout(resolve, 900));
      const worker = WORKERS.find(w => w.id === workerId);
      onScheduled && onScheduled({ hotspotId: hotspot.id, worker: worker?.name, date, notes });
    } catch (e) {
      setStatus("error");
      setErrorMsg("Couldn't reach the scheduling service. Please try again.");
    }
  };

  return (
    <div className={`mt-5 rounded-xl p-4 border ${dark ? "bg-slate-900/50 border-slate-700" : "bg-slate-50 border-slate-200"}`}>
      <div className="flex items-center justify-between mb-3">
        <p className={`text-sm font-bold flex items-center gap-2 ${dark ? "text-white" : "text-slate-900"}`}><Truck size={15} style={{ color: "#0B5D3B" }} /> Schedule Cleaning Crew</p>
        <button type="button" onClick={onCancel} className={`p-1 rounded-lg ${submitting ? "opacity-40 pointer-events-none" : ""} ${dark ? "hover:bg-slate-700 text-slate-400" : "hover:bg-slate-200 text-slate-500"}`}><X size={15} /></button>
      </div>

      <div className="space-y-3">
        <div>
          <label className={labelCls}>Assign Crew Lead</label>
          <select value={workerId} onChange={e => setWorkerId(e.target.value)} className={inputCls} style={lockStyle}>
            <option value="">Select a worker…</option>
            {WORKERS.map(w => <option key={w.id} value={w.id}>{w.name} — {w.ward}{w.status === "Off Duty" ? " (off duty)" : ""}</option>)}
          </select>
        </div>
        <div>
          <label className={labelCls}>Cleaning Date</label>
          <input type="date" min={today} value={date} onChange={e => setDate(e.target.value)} className={inputCls} style={lockStyle} />
        </div>
        <div>
          <label className={labelCls}>Notes (optional)</label>
          <textarea rows={2} value={notes} onChange={e => setNotes(e.target.value)} placeholder="Access instructions, equipment needed…" className={`${inputCls} resize-none`} style={lockStyle} />
        </div>

        {status === "error" && (
          <div className="flex items-start gap-2 text-xs font-medium rounded-lg px-3 py-2" style={{ background: "#FCEBEB", color: "#791F1F" }}>
            <AlertTriangle size={14} className="mt-0.5 flex-shrink-0" /> <span>{errorMsg}</span>
          </div>
        )}

        <div className="flex gap-2">
          <button type="button" onClick={onCancel} className={`flex-1 py-2.5 rounded-xl text-sm font-semibold border ${submitting ? "opacity-50 pointer-events-none" : ""} ${dark ? "border-slate-700 text-slate-300" : "border-slate-200 text-slate-600"}`}>Cancel</button>
          <button type="button" onClick={submit} className="flex-1 py-2.5 rounded-xl text-white text-sm font-semibold flex items-center justify-center gap-2" style={{ background: "#0B5D3B", opacity: submitting ? 0.75 : 1, pointerEvents: submitting ? "none" : "auto", touchAction: "manipulation" }}>
            {submitting
              ? <><span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" /> Scheduling…</>
              : <><CheckCircle2 size={15} /> Confirm</>}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------- Schedule Crew Section (below the map) ---------------------------------- */
// A standalone, always-visible scheduling card placed directly after the map. It has
// its own GVP picker (so it works even without clicking a map marker) and renders the
// inline form — everything in normal document flow, so nothing can block the click.
function ScheduleCrewSection({ selected, onSelect, dark }) {
  const [chosenId, setChosenId] = useState(selected?.id || "");
  const [scheduling, setScheduling] = useState(false);
  const [scheduled, setScheduled] = useState(null);

  // Follow the map/detail selection when it changes.
  useEffect(() => {
    if (selected?.id) { setChosenId(selected.id); setScheduling(false); setScheduled(null); }
  }, [selected?.id]);

  const hotspot = HOTSPOTS.find(h => h.id === chosenId) || null;
  const inputCls = `w-full mt-1 px-3 py-2.5 rounded-xl border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white" : "bg-white border-slate-200 text-slate-900"}`;

  const pick = (id) => {
    setChosenId(id);
    setScheduling(false);
    setScheduled(null);
    const h = HOTSPOTS.find(x => x.id === id);
    if (h && onSelect) onSelect(h); // keep the map + detail panel in sync
  };

  return (
    <Card dark={dark} className="p-5 mt-4">
      <div className="flex items-center gap-2 mb-3">
        <div className="w-8 h-8 rounded-lg flex items-center justify-center" style={{ background: "#0B5D3B" }}>
          <Truck size={16} color="#fff" />
        </div>
        <div>
          <p className={`text-sm font-bold leading-none ${dark ? "text-white" : "text-slate-900"}`}>Schedule Cleaning Crew</p>
          <p className={`text-[11px] mt-0.5 ${dark ? "text-slate-400" : "text-slate-500"}`}>Dispatch a crew to a Garbage Vulnerable Point</p>
        </div>
      </div>

      <label className={`text-xs font-semibold ${dark ? "text-slate-400" : "text-slate-500"}`}>Garbage Vulnerable Point</label>
      <select value={chosenId} onChange={e => pick(e.target.value)} className={inputCls}>
        <option value="">Select a GVP…</option>
        {HOTSPOTS.map(h => <option key={h.id} value={h.id}>{h.name} — {h.ward.split(" - ")[1] || h.ward} ({h.risk})</option>)}
      </select>

      {scheduled && (
        <div className="mt-4 flex items-start gap-2 text-xs font-medium rounded-lg px-3 py-2" style={{ background: "#E1F5EE", color: "#085041" }}>
          <CheckCircle2 size={14} className="mt-0.5 flex-shrink-0" />
          <span>Crew scheduled for {scheduled.date}{scheduled.worker ? ` · ${scheduled.worker}` : ""}.</span>
        </div>
      )}

      {!hotspot ? (
        <p className={`text-xs mt-4 ${dark ? "text-slate-500" : "text-slate-400"}`}>Choose a point above (or click a marker on the map) to schedule a crew.</p>
      ) : scheduling ? (
        <ScheduleCrewForm
          hotspot={hotspot}
          dark={dark}
          onCancel={() => setScheduling(false)}
          onScheduled={(info) => { setScheduled(info); setScheduling(false); }}
        />
      ) : (
        <button
          type="button"
          onClick={() => setScheduling(true)}
          className="w-full mt-4 py-2.5 rounded-xl text-white text-sm font-semibold flex items-center justify-center gap-2 cursor-pointer transition-opacity hover:opacity-95 active:opacity-90"
          style={{ background: "#0B5D3B", touchAction: "manipulation" }}
        >
          <Truck size={15} /> {scheduled ? "Reschedule Crew" : "Schedule Cleaning Crew"}
        </button>
      )}
    </Card>
  );
}

function HotspotDetailPanel({ hotspot, dark, onClose }) {
  const [scheduling, setScheduling] = useState(false);
  const [scheduled, setScheduled] = useState(null);
  // Reset the scheduling UI whenever a different hotspot is selected.
  useEffect(() => { setScheduled(null); setScheduling(false); }, [hotspot?.id]);
  if (!hotspot) return null;
  return (
    <>
    <div style={{ position: "relative", zIndex: 60 }}>
    <Card dark={dark} className="p-5 relative">
      <button onClick={onClose} className={`absolute top-4 right-4 p-1 rounded-lg ${dark ? "hover:bg-slate-700 text-slate-400" : "hover:bg-slate-100 text-slate-400"}`}>
        <X size={16} />
      </button>
      <div className="flex items-center gap-2 mb-1">
        <MapPin size={16} style={{ color: "#0B5D3B" }} />
        <p className={`text-xs font-mono ${dark ? "text-slate-400" : "text-slate-500"}`}>{hotspot.id}</p>
      </div>
      <h3 className={`text-lg font-bold ${dark ? "text-white" : "text-slate-900"}`}>{hotspot.name}</h3>
      <p className={`text-sm ${dark ? "text-slate-400" : "text-slate-500"}`}>{hotspot.ward}</p>
      <div className="flex gap-2 mt-3">
        <RiskBadge risk={hotspot.risk} />
        <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold ${dark ? "bg-slate-700 text-slate-300" : "bg-slate-100 text-slate-600"}`}>
          <AlertTriangle size={12} /> {hotspot.complaints} complaints
        </span>
      </div>

      <div className="grid grid-cols-2 gap-3 mt-4">
        <div className={`rounded-xl p-3 ${dark ? "bg-slate-900/60" : "bg-slate-50"}`}>
          <p className={`text-xs ${dark ? "text-slate-400" : "text-slate-500"}`}>Predicted Recurrence</p>
          <p className="text-2xl font-bold" style={{ color: "#0A4C8C" }}>{hotspot.recurrence}%</p>
        </div>
        <div className={`rounded-xl p-3 ${dark ? "bg-slate-900/60" : "bg-slate-50"}`}>
          <p className={`text-xs ${dark ? "text-slate-400" : "text-slate-500"}`}>Last Cleaned</p>
          <p className={`text-sm font-bold mt-1.5 ${dark ? "text-white" : "text-slate-900"}`}>{hotspot.lastCleaned}</p>
        </div>
      </div>

      <div className="mt-4">
        <p className={`text-xs font-semibold mb-1 ${dark ? "text-slate-400" : "text-slate-500"}`}>AI Root Cause Prediction</p>
        <p className={`text-sm ${dark ? "text-slate-200" : "text-slate-700"}`}>{hotspot.rootCause}</p>
      </div>

      <div className="mt-4">
        <p className={`text-xs font-semibold mb-1.5 ${dark ? "text-slate-400" : "text-slate-500"}`}>Nearby Facilities</p>
        <div className="flex flex-wrap gap-1.5">
          {hotspot.nearby.map((n, i) => (
            <span key={i} className={`text-xs px-2 py-1 rounded-lg ${dark ? "bg-slate-700 text-slate-300" : "bg-slate-100 text-slate-600"}`}>{n}</span>
          ))}
        </div>
      </div>

      {scheduled && (
        <div className="mt-4 flex items-start gap-2 text-xs font-medium rounded-lg px-3 py-2" style={{ background: "#E1F5EE", color: "#085041" }}>
          <CheckCircle2 size={14} className="mt-0.5 flex-shrink-0" />
          <span>Crew scheduled for {scheduled.date}{scheduled.worker ? ` · ${scheduled.worker}` : ""}.</span>
        </div>
      )}

      {scheduling ? (
        <ScheduleCrewForm
          hotspot={hotspot}
          dark={dark}
          onCancel={() => setScheduling(false)}
          onScheduled={(info) => { setScheduled(info); setScheduling(false); }}
        />
      ) : (
        <button
          type="button"
          onClick={() => setScheduling(true)}
          className="w-full mt-5 py-2.5 rounded-xl text-white text-sm font-semibold flex items-center justify-center gap-2 cursor-pointer transition-opacity hover:opacity-95 active:opacity-90"
          style={{ background: "#0B5D3B", touchAction: "manipulation" }}
        >
          <Truck size={15} /> {scheduled ? "Reschedule Crew" : "Schedule Cleaning Crew"}
        </button>
      )}
    </Card>
    </div>
    </>
  );
}

/* ---------------------------------- AUTH MODAL ---------------------------------- */
function AuthModal({ dark, onClose, onAuth, accounts, setAccounts, officers, setOfficers, audience = "citizen" }) {
  const [mode, setMode] = useState("login"); // login | register
  const [showPw, setShowPw] = useState(false);
  const [form, setForm] = useState({ name: "", email: "", phone: "", password: "", uid: "" });
  const [error, setError] = useState("");
  const [created, setCreated] = useState(null);
  const [copied, setCopied] = useState(false);

  // The modal is locked to a single audience — citizens (email) or GVMC officers
  // (UID). The Citizen Portal never exposes officer sign-in, and vice versa.
  const isOfficer = audience === "officer";

  const handleLogin = () => {
    setError("");
    if (isOfficer) {
      // Officers authenticate against the officers database using their UID.
      const uid = form.uid.trim().toUpperCase();
      const match = officers.find(o => String(o.uid).toUpperCase() === uid && o.password === form.password);
      if (!match) { setError("Officer UID or password is incorrect."); return; }
      onAuth(match);
    } else {
      const email = form.email.trim().toLowerCase();
      const match = accounts.find(a => a.email.toLowerCase() === email && a.password === form.password);
      if (!match) { setError("Email or password is incorrect."); return; }
      onAuth(match);
    }
  };

  const handleRegister = () => {
    setError("");
    if (isOfficer) {
      // Officer accounts don't use email — a UID is issued automatically.
      if (!form.name.trim() || !form.phone.trim() || !form.password.trim()) {
        setError("Fill in name, mobile number and password to create an officer account.");
        return;
      }
      const uid = nextOfficerUID(officers);
      const newOfficer = { uid, name: form.name.trim(), phone: form.phone.trim(), password: form.password, role: "admin", designation: "Sanitation Officer" };
      setOfficers([...officers, newOfficer]);
      setCreated(newOfficer);
      return;
    }
    // Citizen registration (email-based).
    if (!form.name.trim() || !form.email.trim() || !form.phone.trim() || !form.password.trim()) {
      setError("Fill in all fields to create an account.");
      return;
    }
    const email = form.email.trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      setError("Enter a valid email address.");
      return;
    }
    if (accounts.some(a => a.email.toLowerCase() === email)) {
      setError("An account with this email already exists. Try logging in instead.");
      return;
    }
    const newAccount = { name: form.name.trim(), email: form.email.trim(), phone: form.phone.trim(), password: form.password, role: "citizen" };
    setAccounts([...accounts, newAccount]);
    setCreated(newAccount);
  };

  const copyCreds = () => {
    const idLine = created.role === "admin" ? `Officer UID: ${created.uid}` : `Email: ${created.email}`;
    navigator.clipboard?.writeText(`${idLine}\nPassword: ${created.password}`);
    setCopied(true);
    setTimeout(() => setCopied(false), 1500);
  };

  if (created) {
    const officerCreated = created.role === "admin";
    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(15,23,42,0.55)", zIndex: 2000 }} onClick={onClose}>
        <div onClick={e => e.stopPropagation()} className={`w-full max-w-sm rounded-2xl p-6 ${dark ? "bg-slate-800" : "bg-white"}`}>
          <div className="w-12 h-12 rounded-full flex items-center justify-center mb-4" style={{ background: "#EAF3DE" }}>
            <CheckCircle2 size={24} style={{ color: "#3B6D11" }} />
          </div>
          <h3 className={`text-lg font-bold ${dark ? "text-white" : "text-slate-900"}`}>{officerCreated ? "Officer account created" : "Account created"}</h3>
          <p className={`text-sm mt-1 mb-4 ${dark ? "text-slate-400" : "text-slate-500"}`}>{officerCreated ? "Your Officer UID is your login — save it somewhere safe. You'll sign in with the UID, not an email." : "Here are your login credentials — save them somewhere safe. You'll also need them to sign back in."}</p>
          <div className={`rounded-xl p-4 space-y-2 text-sm ${dark ? "bg-slate-900/60" : "bg-slate-50"}`}>
            <div className="flex justify-between"><span className={dark ? "text-slate-400" : "text-slate-500"}>Name</span><span className={`font-semibold ${dark ? "text-white" : "text-slate-900"}`}>{created.name}</span></div>
            {officerCreated ? (
              <div className="flex justify-between"><span className={dark ? "text-slate-400" : "text-slate-500"}>Officer UID</span><span className={`font-mono font-semibold ${dark ? "text-white" : "text-slate-900"}`}>{created.uid}</span></div>
            ) : (
              <div className="flex justify-between"><span className={dark ? "text-slate-400" : "text-slate-500"}>Email</span><span className={`font-mono font-semibold ${dark ? "text-white" : "text-slate-900"}`}>{created.email}</span></div>
            )}
            <div className="flex justify-between"><span className={dark ? "text-slate-400" : "text-slate-500"}>Password</span><span className={`font-mono font-semibold ${dark ? "text-white" : "text-slate-900"}`}>{created.password}</span></div>
            <div className="flex justify-between"><span className={dark ? "text-slate-400" : "text-slate-500"}>Role</span><span className={`font-semibold ${dark ? "text-white" : "text-slate-900"}`}>{officerCreated ? "GVMC Officer" : "Citizen"}</span></div>
          </div>
          <button onClick={copyCreds} className={`w-full mt-3 py-2.5 rounded-xl text-sm font-semibold border-2 flex items-center justify-center gap-2 ${dark ? "border-slate-700 text-slate-200" : "border-slate-200 text-slate-700"}`}>
            {copied ? <Check size={14} /> : <Copy size={14} />} {copied ? "Copied" : "Copy credentials"}
          </button>
          <button onClick={() => onAuth(created)} className="w-full mt-2 py-2.5 rounded-xl text-white text-sm font-semibold" style={{ background: "#0B5D3B" }}>
            Continue to portal
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(15,23,42,0.55)", zIndex: 2000 }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} className={`w-full max-w-sm rounded-2xl p-6 relative ${dark ? "bg-slate-800" : "bg-white"}`}>
        <button onClick={onClose} className={`absolute top-4 right-4 p-1 rounded-lg ${dark ? "hover:bg-slate-700 text-slate-400" : "hover:bg-slate-100 text-slate-400"}`}><X size={16} /></button>

        <div className="flex items-center gap-2.5 mb-4">
          <div className="w-9 h-9 rounded-xl flex items-center justify-center" style={{ background: isOfficer ? "#0A4C8C" : "#0B5D3B" }}>
            {isOfficer ? <Shield size={17} color="#fff" /> : <Users size={17} color="#fff" />}
          </div>
          <div>
            <p className={`text-sm font-bold leading-none ${dark ? "text-white" : "text-slate-900"}`}>{isOfficer ? "GVMC Officer Access" : "Citizen Portal"}</p>
            <p className={`text-[11px] mt-0.5 ${dark ? "text-slate-400" : "text-slate-500"}`}>{isOfficer ? "Authorised municipal staff only" : "Report & track garbage complaints"}</p>
          </div>
        </div>

        <div className={`flex gap-1 p-1 rounded-xl mb-5 w-fit ${dark ? "bg-slate-900" : "bg-slate-100"}`}>
          <button onClick={() => { setMode("login"); setError(""); }} className={`px-4 py-1.5 rounded-lg text-sm font-semibold ${mode === "login" ? "bg-white shadow text-emerald-700" : dark ? "text-slate-400" : "text-slate-500"}`}>Login</button>
          <button onClick={() => { setMode("register"); setError(""); }} className={`px-4 py-1.5 rounded-lg text-sm font-semibold ${mode === "register" ? "bg-white shadow text-emerald-700" : dark ? "text-slate-400" : "text-slate-500"}`}>Register</button>
        </div>

        <h3 className={`text-lg font-bold ${dark ? "text-white" : "text-slate-900"}`}>{mode === "login" ? "Welcome back" : "Create your account"}</h3>
        <p className={`text-xs mt-1 mb-4 ${dark ? "text-slate-400" : "text-slate-500"}`}>{isOfficer ? "GVMC officers sign in with their Officer UID." : "Citizens sign in with their email address."}</p>

        <div className="space-y-3">
          {mode === "register" && (
            <div className="relative">
              <Users size={15} className={`absolute left-3 top-1/2 -translate-y-1/2 ${dark ? "text-slate-500" : "text-slate-400"}`} />
              <input value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="Full name" className={`w-full pl-9 pr-3 py-2.5 rounded-xl border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white placeholder:text-slate-500" : "bg-white border-slate-200 text-slate-900 placeholder:text-slate-400"}`} />
            </div>
          )}

          {/* Identifier field: Officer UID (login only) OR Email (citizens) */}
          {isOfficer ? (
            mode === "login" ? (
              <div className="relative">
                <Shield size={15} className={`absolute left-3 top-1/2 -translate-y-1/2 ${dark ? "text-slate-500" : "text-slate-400"}`} />
                <input value={form.uid} onChange={e => setForm({ ...form, uid: e.target.value })} placeholder="Officer UID (e.g. GVMC-OFF-001)" className={`w-full pl-9 pr-3 py-2.5 rounded-xl border text-sm font-mono ${dark ? "bg-slate-900 border-slate-700 text-white placeholder:text-slate-500" : "bg-white border-slate-200 text-slate-900 placeholder:text-slate-400"}`} />
              </div>
            ) : (
              <div className={`rounded-xl p-3 text-xs flex items-start gap-2 ${dark ? "bg-slate-900/60 text-slate-400" : "bg-blue-50 text-blue-700"}`}>
                <Info size={14} className="mt-0.5 flex-shrink-0" />
                <span>A unique Officer UID will be generated automatically when you create this account.</span>
              </div>
            )
          ) : (
            <div className="relative">
              <Mail size={15} className={`absolute left-3 top-1/2 -translate-y-1/2 ${dark ? "text-slate-500" : "text-slate-400"}`} />
              <input value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} placeholder="Email address" className={`w-full pl-9 pr-3 py-2.5 rounded-xl border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white placeholder:text-slate-500" : "bg-white border-slate-200 text-slate-900 placeholder:text-slate-400"}`} />
            </div>
          )}

          {mode === "register" && (
            <div className="relative">
              <Phone size={15} className={`absolute left-3 top-1/2 -translate-y-1/2 ${dark ? "text-slate-500" : "text-slate-400"}`} />
              <input value={form.phone} onChange={e => setForm({ ...form, phone: e.target.value })} placeholder="Mobile number" className={`w-full pl-9 pr-3 py-2.5 rounded-xl border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white placeholder:text-slate-500" : "bg-white border-slate-200 text-slate-900 placeholder:text-slate-400"}`} />
            </div>
          )}

          <div className="relative">
            <Lock size={15} className={`absolute left-3 top-1/2 -translate-y-1/2 ${dark ? "text-slate-500" : "text-slate-400"}`} />
            <input type={showPw ? "text" : "password"} value={form.password} onChange={e => setForm({ ...form, password: e.target.value })} placeholder="Password" className={`w-full pl-9 pr-9 py-2.5 rounded-xl border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white placeholder:text-slate-500" : "bg-white border-slate-200 text-slate-900 placeholder:text-slate-400"}`} />
            <button onClick={() => setShowPw(!showPw)} className={`absolute right-3 top-1/2 -translate-y-1/2 ${dark ? "text-slate-500" : "text-slate-400"}`}>{showPw ? <EyeOff size={15} /> : <Eye size={15} />}</button>
          </div>

          {error && <p className="text-xs font-medium" style={{ color: "#D6432B" }}>{error}</p>}

          <button onClick={mode === "login" ? handleLogin : handleRegister} className="w-full py-2.5 rounded-xl text-white text-sm font-semibold" style={{ background: "#0B5D3B" }}>
            {mode === "login" ? (isOfficer ? "Login as Officer" : "Login") : "Create account"}
          </button>
        </div>

        {mode === "login" && (
          <div className={`mt-4 rounded-xl p-3 text-xs ${dark ? "bg-slate-900/60 text-slate-400" : "bg-slate-50 text-slate-500"}`}>
            <p className="font-semibold mb-1">Demo credentials</p>
            {isOfficer ? <p>Officer — UID GVMC-OFF-001 / admin123</p> : <p>Citizen — citizen@gvmc.gov.in / citizen123</p>}
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------------------------------- AUTH GATE ---------------------------------- */
function AuthGate({ dark, onLogin, label }) {
  return (
    <div className="max-w-md mx-auto px-4 py-24 text-center">
      <div className="w-14 h-14 rounded-2xl flex items-center justify-center mx-auto mb-4" style={{ background: dark ? "#0f2a20" : "#eaf3ee" }}>
        <Lock size={22} style={{ color: "#0B5D3B" }} />
      </div>
      <h2 className={`text-lg font-bold ${dark ? "text-white" : "text-slate-900"}`}>Sign in to continue</h2>
      <p className={`text-sm mt-1.5 mb-6 ${dark ? "text-slate-400" : "text-slate-500"}`}>You need an account to access {label}.</p>
      <button onClick={onLogin} className="px-5 py-2.5 rounded-xl text-white text-sm font-semibold" style={{ background: "#0B5D3B" }}>Login or create an account</button>
    </div>
  );
}

/* ---------------------------------- ACCESS DENIED ---------------------------------- */
function AccessDenied({ dark, onSwitch }) {
  return (
    <div className="max-w-md mx-auto px-4 py-24 text-center">
      <div className="w-14 h-14 rounded-2xl flex items-center justify-center mx-auto mb-4" style={{ background: dark ? "#3a2410" : "#fdf1e3" }}>
        <Shield size={22} style={{ color: "#B4540A" }} />
      </div>
      <h2 className={`text-lg font-bold ${dark ? "text-white" : "text-slate-900"}`}>Access restricted</h2>
      <p className={`text-sm mt-1.5 mb-6 ${dark ? "text-slate-400" : "text-slate-500"}`}>You're signed in as a citizen. The Admin Dashboard is available only to GVMC officers. To continue, sign out and log in with an officer account.</p>
      <button onClick={onSwitch} className="px-5 py-2.5 rounded-xl text-white text-sm font-semibold" style={{ background: "#0B5D3B" }}>Log out &amp; switch account</button>
    </div>
  );
}

/* ---------------------------------- NAVBAR ---------------------------------- */
function Navbar({ view, setView, dark, setDark, lang, setLang, role, setRole, t, mobileOpen, setMobileOpen, user, setShowAuth, onLogout }) {
  const navItems = [
    { key: "home", label: t.home, icon: Home },
    { key: "citizen", label: t.citizen, icon: Users },
    { key: "admin", label: t.admin, icon: LayoutDashboard },
    { key: "ai", label: t.ai, icon: Brain },
  ];
  return (
    <nav className={`sticky top-0 z-40 border-b backdrop-blur-md ${dark ? "bg-slate-900/90 border-slate-800" : "bg-white/90 border-slate-200"}`}>
      <div className="max-w-7xl mx-auto px-4 sm:px-6 h-16 flex items-center justify-between">
        <div className="flex items-center gap-2.5 cursor-pointer" onClick={() => setView("home")}>
          <div className="w-12 h-12 flex-shrink-0">
            <svg viewBox="0 0 48 48" className="w-12 h-12" role="img" aria-label="Swachha GVMC logo" xmlns="http://www.w3.org/2000/svg">
              <defs>
                <linearGradient id="gvmcLogoBg" x1="0" y1="0" x2="1" y2="1">
                  <stop offset="0%" stopColor="#0F766E" />
                  <stop offset="58%" stopColor="#15803D" />
                  <stop offset="100%" stopColor="#2563EB" />
                </linearGradient>
              </defs>
              {/* Rounded-square badge */}
              <rect x="0" y="0" width="48" height="48" rx="13" fill="url(#gvmcLogoBg)" />
              {/* AI / network nodes — intelligent monitoring */}
              <g opacity="0.9">
                <line x1="9" y1="11" x2="17" y2="18" stroke="#22C55E" strokeWidth="1.2" />
                <line x1="39" y1="13" x2="31" y2="18" stroke="#22C55E" strokeWidth="1.2" />
                <line x1="10" y1="37" x2="20" y2="31" stroke="#22C55E" strokeWidth="1.2" />
                <circle cx="9" cy="11" r="1.8" fill="#22C55E" />
                <circle cx="39" cy="13" r="1.8" fill="#22C55E" />
                <circle cx="10" cy="37" r="1.8" fill="#22C55E" />
              </g>
              {/* Location pin — the garbage vulnerable point */}
              <path d="M24 8c-6.6 0-12 5.2-12 11.8 0 8 10 17.1 11.4 18.4a0.9 0.9 0 0 0 1.2 0C26 36.9 36 27.8 36 19.8 36 13.2 30.6 8 24 8z" fill="#ffffff" />
              {/* Leaf — cleanliness & sustainability */}
              <path d="M18.6 22.6c-0.5-6 4.3-10.2 11-10.3 0.6 6-4.2 10.3-11 10.3z" fill="#22C55E" />
              <path d="M18.8 22.4C22 20 25.4 16.8 28.8 13.1" stroke="#ffffff" strokeWidth="1.3" strokeLinecap="round" fill="none" />
            </svg>
          </div>
          <div>
            <p className={`text-sm font-bold leading-none ${dark ? "text-white" : "text-slate-900"}`}>{t.appName}</p>
            <p className={`text-[10px] ${dark ? "text-slate-400" : "text-slate-500"}`}>{t.tagline}</p>
          </div>
        </div>

        <div className="hidden lg:flex items-center gap-1">
          {navItems.map((item) => (
            <button
              key={item.key}
              onClick={() => setView(item.key)}
              className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium transition-colors ${
                view === item.key
                  ? dark ? "bg-emerald-900/40 text-emerald-300" : "bg-emerald-50 text-emerald-700"
                  : dark ? "text-slate-300 hover:bg-slate-800" : "text-slate-600 hover:bg-slate-100"
              }`}
            >
              <item.icon size={15} /> {item.label}
            </button>
          ))}
        </div>

        <div className="hidden md:flex items-center gap-2">
          <button onClick={() => setLang(lang === "en" ? "te" : "en")} className={`p-2 rounded-lg text-xs font-bold ${dark ? "hover:bg-slate-800 text-slate-300" : "hover:bg-slate-100 text-slate-600"}`} title="Switch language">
            <Globe size={16} />
          </button>
          <button onClick={() => setDark(!dark)} className={`p-2 rounded-lg ${dark ? "hover:bg-slate-800 text-amber-300" : "hover:bg-slate-100 text-slate-600"}`}>
            {dark ? <Sun size={16} /> : <Moon size={16} />}
          </button>
          <div className={`flex items-center rounded-lg p-0.5 text-xs font-semibold ${dark ? "bg-slate-800" : "bg-slate-100"}`}>
            <button type="button" onClick={() => setView("citizen")} aria-label="Go to Citizen Portal" className={`px-2.5 py-1.5 rounded-md ${view === "citizen" ? "bg-white shadow text-emerald-700" : dark ? "text-slate-400 hover:text-slate-200" : "text-slate-500 hover:text-slate-700"}`}>Citizen</button>
            <button type="button" onClick={() => setView("admin")} aria-label="Go to Officer Dashboard" className={`px-2.5 py-1.5 rounded-md ${view === "admin" ? "bg-white shadow text-blue-700" : dark ? "text-slate-400 hover:text-slate-200" : "text-slate-500 hover:text-slate-700"}`}>Officer</button>
          </div>
          {user ? (
            <div className="flex items-center gap-2">
              <div className="w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold text-white" style={{ background: user.role === "admin" ? "#0A4C8C" : "#0B5D3B" }}>
                {user.name.split(" ").map(n => n[0]).join("").slice(0, 2)}
              </div>
              <button onClick={onLogout} className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold ${dark ? "text-slate-300 hover:bg-slate-800" : "text-slate-600 hover:bg-slate-100"}`}>
                <LogOut size={14} /> {t.logout}
              </button>
            </div>
          ) : (
            <button onClick={() => setShowAuth(true)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: "#0B5D3B" }}>
              <LogIn size={14} /> {t.login}
            </button>
          )}
        </div>

        <button className="lg:hidden p-2" onClick={() => setMobileOpen(!mobileOpen)}>
          <Menu size={22} className={dark ? "text-white" : "text-slate-900"} />
        </button>
      </div>

      {mobileOpen && (
        <div className={`lg:hidden border-t px-4 py-3 space-y-1 ${dark ? "border-slate-800 bg-slate-900" : "border-slate-200 bg-white"}`}>
          {navItems.map((item) => (
            <button key={item.key} onClick={() => { setView(item.key); setMobileOpen(false); }}
              className={`flex items-center gap-2 w-full px-3 py-2.5 rounded-lg text-sm font-medium ${view === item.key ? "bg-emerald-50 text-emerald-700" : dark ? "text-slate-300" : "text-slate-600"}`}>
              <item.icon size={16} /> {item.label}
            </button>
          ))}
          <div className="flex items-center gap-2 pt-2">
            <button onClick={() => setLang(lang === "en" ? "te" : "en")} className="flex-1 py-2 rounded-lg text-xs font-semibold border">EN / తె</button>
            <button onClick={() => setDark(!dark)} className="flex-1 py-2 rounded-lg text-xs font-semibold border">{dark ? "Light" : "Dark"} Mode</button>
          </div>
          {user ? (
            <button onClick={onLogout} className="w-full mt-2 py-2 rounded-lg text-xs font-semibold border flex items-center justify-center gap-1.5"><LogOut size={13} /> {t.logout}</button>
          ) : (
            <button onClick={() => { setShowAuth(true); setMobileOpen(false); }} className="w-full mt-2 py-2 rounded-lg text-xs font-semibold text-white flex items-center justify-center gap-1.5" style={{ background: "#0B5D3B" }}><LogIn size={13} /> {t.login}</button>
          )}
        </div>
      )}
    </nav>
  );
}

/* ---------------------------------- HOME PAGE ---------------------------------- */
/* ---------------------------------- AI Analytics Dashboard ---------------------------------- */
// Simple linear-regression forecast for next-month complaints (the "AI predicted" line).
function forecastNext(values) {
  const n = values.length;
  if (!n) return 0;
  const xs = values.map((_, i) => i);
  const meanX = xs.reduce((a, b) => a + b, 0) / n;
  const meanY = values.reduce((a, b) => a + b, 0) / n;
  let num = 0, den = 0;
  for (let i = 0; i < n; i++) { num += (xs[i] - meanX) * (values[i] - meanY); den += (xs[i] - meanX) ** 2; }
  const slope = den ? num / den : 0;
  const intercept = meanY - slope * meanX;
  return Math.max(0, Math.round(intercept + slope * n));
}

function KpiCard({ icon: Icon, label, value, sub, accent, dark }) {
  return (
    <div className={`rounded-2xl p-4 backdrop-blur shadow-sm ${dark ? "border border-slate-700/60 bg-slate-800/50" : "border border-white/70 bg-white/70"}`}>
      <div className="flex items-center justify-between">
        <div style={{ background: accent + "1f" }} className="w-9 h-9 rounded-xl flex items-center justify-center">
          <Icon size={17} style={{ color: accent }} />
        </div>
      </div>
      <p className={`text-2xl font-extrabold mt-2 ${dark ? "text-white" : "text-slate-900"}`}>{value}</p>
      <p className={`text-[11px] font-semibold uppercase tracking-wide ${dark ? "text-slate-400" : "text-slate-500"}`}>{label}</p>
      {sub && <p className="text-[11px] mt-0.5" style={{ color: accent }}>{sub}</p>}
    </div>
  );
}

function ProgressRing({ pct, dark, color = "#0B5D3B", caption }) {
  const r = 52, C = 2 * Math.PI * r, off = C * (1 - Math.max(0, Math.min(100, pct)) / 100);
  return (
    <svg viewBox="0 0 140 140" className="w-36 h-36">
      <circle cx="70" cy="70" r={r} fill="none" stroke={dark ? "#334155" : "#e5edf5"} strokeWidth="12" />
      <circle cx="70" cy="70" r={r} fill="none" stroke={color} strokeWidth="12" strokeLinecap="round"
        strokeDasharray={C} strokeDashoffset={off} transform="rotate(-90 70 70)" style={{ transition: "stroke-dashoffset 0.7s ease" }} />
      <text x="70" y="66" textAnchor="middle" fontSize="28" fontWeight="800" fill={dark ? "#fff" : "#0f172a"}>{pct}%</text>
      <text x="70" y="88" textAnchor="middle" fontSize="10" fill={dark ? "#94a3b8" : "#64748b"}>{caption}</text>
    </svg>
  );
}

function AnalyticsDashboard({ complaints, dark, selected, setSelected }) {
  const RISK = ["High", "Medium", "Low"];
  const [fWard, setFWard] = useState("All");
  const [fStatus, setFStatus] = useState("All");
  const [fType, setFType] = useState("All");
  const [fRisk, setFRisk] = useState("All");
  const [fFrom, setFFrom] = useState("");
  const [fTo, setFTo] = useState("");
  const [timelineId, setTimelineId] = useState("");

  const fc = useMemo(() => complaints.filter(c =>
    (fWard === "All" || c.ward === fWard) &&
    (fStatus === "All" || c.status === fStatus) &&
    (fType === "All" || c.type === fType) &&
    (!fFrom || c.date >= fFrom) &&
    (!fTo || c.date <= fTo)
  ), [complaints, fWard, fStatus, fType, fFrom, fTo]);

  const fh = useMemo(() => HOTSPOTS.filter(h =>
    (fWard === "All" || h.ward === fWard) &&
    (fRisk === "All" || h.risk === fRisk)
  ), [fWard, fRisk]);

  const resolvedSet = ["Cleaned", "Closed"];
  const total = fc.length;
  const closed = fc.filter(c => resolvedSet.includes(c.status)).length;
  const active = total - closed;
  const resolutionRate = total ? Math.round((closed / total) * 100) : 0;

  // Avg cleaning time from status history (Submitted -> Cleaned), else a sensible default.
  const avgCleanDays = useMemo(() => {
    const spans = [];
    fc.forEach(c => {
      const h = c.statusHistory || [];
      const start = h.find(x => x.status === "Submitted");
      const done = h.find(x => x.status === "Cleaned" || x.status === "Closed");
      if (start && done) spans.push((new Date(done.at) - new Date(start.at)) / 86400000);
    });
    if (!spans.length) return 2.4;
    return Math.round((spans.reduce((a, b) => a + b, 0) / spans.length) * 10) / 10;
  }, [fc]);

  const statusDist = useMemo(() => STATUSES.map(s => ({ name: s, value: fc.filter(c => c.status === s).length })).filter(x => x.value > 0), [fc]);

  const wardAvgRes = (ward) => { let h = 0; for (const ch of ward) h = (h * 31 + ch.charCodeAt(0)) % 100000; return Math.round((1.5 + (h % 250) / 100) * 10) / 10; };
  const wardStats = useMemo(() => WARDS.map(w => {
    const short = w.split(" - ")[1] || w;
    return { ward: short, full: w, complaints: fc.filter(c => c.ward === w).length, avgRes: wardAvgRes(short) };
  }).sort((a, b) => b.complaints - a.complaints).slice(0, 10), [fc]);
  const maxWard = wardStats.length ? Math.max(...wardStats.map(w => w.complaints)) : 0;

  const riskMix = useMemo(() => {
    const row = { name: "GVP Risk Mix" };
    RISK.forEach(r => { row[r] = fh.filter(h => h.risk === r).length; });
    return [row];
  }, [fh]);
  const riskCounts = { High: fh.filter(h => h.risk === "High").length, Medium: fh.filter(h => h.risk === "Medium").length, Low: fh.filter(h => h.risk === "Low").length };

  const trendData = useMemo(() => {
    const pred = forecastNext(TREND_DATA.map(d => d.complaints));
    const rows = TREND_DATA.map((d, i) => ({ ...d, predicted: i === TREND_DATA.length - 1 ? d.complaints : null }));
    rows.push({ month: "Aug (AI)", complaints: null, cleaned: null, predicted: pred });
    return rows;
  }, []);

  const glass = `rounded-2xl p-5 backdrop-blur shadow-sm ${dark ? "border border-slate-700/60 bg-slate-800/50" : "border border-white/70 bg-white/70"}`;
  const selCls = `px-3 py-2 rounded-xl border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white" : "bg-white border-slate-200 text-slate-900"}`;

  // Black-spot detail helpers
  const nearbyOf = (h, kw) => (h.nearby || []).find(n => n.toLowerCase().includes(kw)) || "—";
  const workerOf = (h) => (WORKERS.find(w => h.ward.startsWith(w.ward)) || {}).name || "Unassigned";
  const gvpStatus = (h) => h.recurrence > 70 ? "Active · High recurrence" : h.recurrence > 40 ? "Monitoring" : "Stable";

  const tl = complaints.find(c => c.id === timelineId) || fc[0] || complaints[0];
  const tlHist = {};
  (tl?.statusHistory || []).forEach(x => { tlHist[x.status] = x.at; });
  const tlIdx = tl ? Math.max(0, TRACK_STAGES.indexOf(tl.status)) : 0;
  const fmt = (iso) => { try { return new Date(iso).toLocaleString("en-IN", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }); } catch { return ""; } };

  const PIE = ["#639922", "#EF9F27", "#378ADD", "#7C3AED", "#1D9E75", "#888780"];

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 pb-20">
      <div className="rounded-3xl p-4 sm:p-6" style={{ background: dark ? "linear-gradient(135deg,#0a1f16,#0a1830)" : "linear-gradient(135deg,#eef7f1,#eaf2fb)" }}>
        {/* Header + filters */}
        <div className="flex items-start justify-between flex-wrap gap-3 mb-5">
          <div>
            <p className="text-xs font-bold tracking-widest uppercase" style={{ color: "#0B5D3B" }}>AI Analytics</p>
            <h2 className={`text-2xl font-extrabold ${dark ? "text-white" : "text-slate-900"}`}>GVP Prediction &amp; Prevention Dashboard</h2>
            <p className={`text-sm mt-1 ${dark ? "text-slate-400" : "text-slate-500"}`}>Live insights across complaints, wards, risk and resolution performance.</p>
          </div>
        </div>

        <div className={`${glass} mb-5`}>
          <div className="flex items-center gap-2 mb-3">
            <Filter size={15} style={{ color: "#0A4C8C" }} />
            <p className={`text-sm font-bold ${dark ? "text-white" : "text-slate-900"}`}>Filters</p>
            <button onClick={() => { setFWard("All"); setFStatus("All"); setFType("All"); setFRisk("All"); setFFrom(""); setFTo(""); }} className={`ml-auto text-xs font-semibold ${dark ? "text-slate-400 hover:text-slate-200" : "text-slate-500 hover:text-slate-700"}`}>Reset</button>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-2">
            <input type="date" value={fFrom} onChange={e => setFFrom(e.target.value)} className={selCls} aria-label="From date" />
            <input type="date" value={fTo} onChange={e => setFTo(e.target.value)} className={selCls} aria-label="To date" />
            <select value={fWard} onChange={e => setFWard(e.target.value)} className={selCls}><option value="All">All Wards</option>{WARDS.map(w => <option key={w} value={w}>{w.split(" - ")[1] || w}</option>)}</select>
            <select value={fStatus} onChange={e => setFStatus(e.target.value)} className={selCls}><option value="All">All Status</option>{STATUSES.map(s => <option key={s}>{s}</option>)}</select>
            <select value={fType} onChange={e => setFType(e.target.value)} className={selCls}><option value="All">All Types</option>{GARBAGE_TYPES.map(g => <option key={g}>{g}</option>)}</select>
            <select value={fRisk} onChange={e => setFRisk(e.target.value)} className={selCls}><option value="All">All Risk</option>{RISK.map(r => <option key={r}>{r}</option>)}</select>
          </div>
        </div>

        {/* KPI cards */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-5">
          <KpiCard icon={ClipboardList} label="Total Complaints" value={total} accent="#0A4C8C" dark={dark} />
          <KpiCard icon={Clock} label="Active Complaints" value={active} sub="in progress" accent="#E8A93B" dark={dark} />
          <KpiCard icon={CheckCircle2} label="Closed Complaints" value={closed} sub={`${resolutionRate}% resolved`} accent="#0B5D3B" dark={dark} />
          <KpiCard icon={AlertTriangle} label="High-Risk Areas" value={riskCounts.High} accent="#D6432B" dark={dark} />
          <KpiCard icon={TrendingUp} label="Predicted Recurring" value={fh.filter(h => h.recurrence > 70).length} sub=">70% recurrence" accent="#7C3AED" dark={dark} />
          <KpiCard icon={Truck} label="Avg Cleaning Time" value={`${avgCleanDays}d`} accent="#1D9E75" dark={dark} />
          <KpiCard icon={Brain} label="AI Accuracy" value="94.2%" sub="XGBoost" accent="#0A4C8C" dark={dark} />
          <KpiCard icon={Star} label="Citizen Satisfaction" value="91%" accent="#EF9F27" dark={dark} />
        </div>

        {/* Trend + Status */}
        <div className="grid lg:grid-cols-3 gap-4 mb-5">
          <div className={`${glass} lg:col-span-2`}>
            <p className={`text-sm font-bold ${dark ? "text-white" : "text-slate-900"}`}>Monthly Complaint &amp; Resolution Trend</p>
            <p className={`text-xs mb-3 ${dark ? "text-slate-400" : "text-slate-500"}`}>Registered vs resolved, with AI-forecast for next month (dashed). Drag the slider to zoom.</p>
            <div style={{ width: "100%", height: 280 }}>
              <ResponsiveContainer>
                <LineChart data={trendData} margin={{ top: 5, right: 10, left: -10, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#334155" : "#e2e8f0"} vertical={false} />
                  <XAxis dataKey="month" tick={{ fontSize: 11, fill: dark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: dark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ borderRadius: 12, border: "none", fontSize: 12 }} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Line type="monotone" dataKey="complaints" name="Registered" stroke="#0A4C8C" strokeWidth={2.5} dot={{ r: 3 }} connectNulls />
                  <Line type="monotone" dataKey="cleaned" name="Resolved" stroke="#0B5D3B" strokeWidth={2.5} dot={{ r: 3 }} connectNulls />
                  <Line type="monotone" dataKey="predicted" name="AI Predicted" stroke="#EA8C1B" strokeWidth={2.5} strokeDasharray="6 5" dot={{ r: 4 }} connectNulls />
                  <Brush dataKey="month" height={18} stroke={dark ? "#475569" : "#94a3b8"} travellerWidth={8} />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className={glass}>
            <p className={`text-sm font-bold ${dark ? "text-white" : "text-slate-900"}`}>Complaint Status Distribution</p>
            <p className={`text-xs mb-1 ${dark ? "text-slate-400" : "text-slate-500"}`}>Total {total} complaints</p>
            <div style={{ width: "100%", height: 200 }}>
              <ResponsiveContainer>
                <PieChart>
                  <Pie data={statusDist} dataKey="value" nameKey="name" innerRadius={48} outerRadius={80} paddingAngle={2}>
                    {statusDist.map((e, i) => <Cell key={i} fill={(STATUS_COLORS[e.name] || {}).dot || PIE[i % PIE.length]} />)}
                  </Pie>
                  <Tooltip formatter={(v, n) => [`${v} (${total ? Math.round((v / total) * 100) : 0}%)`, n]} contentStyle={{ borderRadius: 12, border: "none", fontSize: 12 }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div className="flex flex-wrap gap-x-3 gap-y-1 mt-1">
              {statusDist.map(e => (
                <span key={e.name} className="flex items-center gap-1 text-[11px]" style={{ color: dark ? "#cbd5e1" : "#475569" }}>
                  <span className="w-2 h-2 rounded-full" style={{ background: (STATUS_COLORS[e.name] || {}).dot }}></span>{e.name} {total ? Math.round((e.value / total) * 100) : 0}%
                </span>
              ))}
            </div>
          </div>
        </div>

        {/* Ward + Risk + Resolution */}
        <div className="grid lg:grid-cols-3 gap-4 mb-5">
          <div className={`${glass} lg:col-span-1`}>
            <p className={`text-sm font-bold ${dark ? "text-white" : "text-slate-900"}`}>Ward-wise Complaints</p>
            <p className={`text-xs mb-3 ${dark ? "text-slate-400" : "text-slate-500"}`}>Top wards — highest highlighted</p>
            <div style={{ width: "100%", height: 260 }}>
              <ResponsiveContainer>
                <BarChart data={wardStats} layout="vertical" margin={{ left: 6, right: 12 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#334155" : "#e2e8f0"} horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 10 }} axisLine={false} tickLine={false} allowDecimals={false} />
                  <YAxis type="category" dataKey="ward" width={92} tick={{ fontSize: 10, fill: dark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} />
                  <Tooltip formatter={(v, n, p) => [`${v} complaints · ~${p.payload.avgRes}d avg`, "Ward"]} contentStyle={{ borderRadius: 12, border: "none", fontSize: 12 }} />
                  <Bar dataKey="complaints" radius={[0, 6, 6, 0]} barSize={14}>
                    {wardStats.map((w, i) => <Cell key={i} fill={w.complaints === maxWard && maxWard > 0 ? "#D6432B" : "#0A4C8C"} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className={glass}>
            <p className={`text-sm font-bold ${dark ? "text-white" : "text-slate-900"}`}>AI Risk Level Distribution</p>
            <p className={`text-xs mb-3 ${dark ? "text-slate-400" : "text-slate-500"}`}>Garbage Vulnerable Points by risk tier</p>
            <div style={{ width: "100%", height: 180 }}>
              <ResponsiveContainer>
                <BarChart data={riskMix} margin={{ left: -20 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#334155" : "#e2e8f0"} vertical={false} />
                  <XAxis dataKey="name" tick={{ fontSize: 11, fill: dark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize: 10 }} axisLine={false} tickLine={false} allowDecimals={false} />
                  <Tooltip contentStyle={{ borderRadius: 12, border: "none", fontSize: 12 }} />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                  <Bar dataKey="High" stackId="r" fill="#E24B4A" radius={[0, 0, 0, 0]} barSize={70} />
                  <Bar dataKey="Medium" stackId="r" fill="#EF9F27" barSize={70} />
                  <Bar dataKey="Low" stackId="r" fill="#639922" radius={[6, 6, 0, 0]} barSize={70} />
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="flex justify-around mt-1 text-center">
              <div><p className="text-lg font-extrabold" style={{ color: "#E24B4A" }}>{riskCounts.High}</p><p className="text-[10px]" style={{ color: dark ? "#94a3b8" : "#64748b" }}>High</p></div>
              <div><p className="text-lg font-extrabold" style={{ color: "#EF9F27" }}>{riskCounts.Medium}</p><p className="text-[10px]" style={{ color: dark ? "#94a3b8" : "#64748b" }}>Medium</p></div>
              <div><p className="text-lg font-extrabold" style={{ color: "#639922" }}>{riskCounts.Low}</p><p className="text-[10px]" style={{ color: dark ? "#94a3b8" : "#64748b" }}>Low</p></div>
            </div>
          </div>

          <div className={`${glass} flex flex-col items-center justify-center`}>
            <p className={`text-sm font-bold self-start ${dark ? "text-white" : "text-slate-900"}`}>Resolution Performance</p>
            <ProgressRing pct={resolutionRate} dark={dark} caption="Resolved" color="#0B5D3B" />
            <div className="grid grid-cols-2 gap-2 w-full mt-2">
              <div className={`rounded-xl p-2 text-center ${dark ? "bg-slate-900/50" : "bg-slate-50"}`}><p className="text-base font-bold" style={{ color: "#0A4C8C" }}>{avgCleanDays}d</p><p className="text-[10px]" style={{ color: dark ? "#94a3b8" : "#64748b" }}>Avg time</p></div>
              <div className={`rounded-xl p-2 text-center ${dark ? "bg-slate-900/50" : "bg-slate-50"}`}><p className="text-base font-bold" style={{ color: "#E8A93B" }}>{active}</p><p className="text-[10px]" style={{ color: dark ? "#94a3b8" : "#64748b" }}>Pending</p></div>
              <div className={`rounded-xl p-2 text-center ${dark ? "bg-slate-900/50" : "bg-slate-50"} col-span-2`}><p className="text-base font-bold" style={{ color: "#0B5D3B" }}>{closed}</p><p className="text-[10px]" style={{ color: dark ? "#94a3b8" : "#64748b" }}>Completed</p></div>
            </div>
          </div>
        </div>

        {/* Heatmap + Black spot detail */}
        <div className="grid lg:grid-cols-3 gap-4 mb-5">
          <div className={`${glass} lg:col-span-2`}>
            <p className={`text-sm font-bold flex items-center gap-2 ${dark ? "text-white" : "text-slate-900"}`}><MapIcon size={15} style={{ color: "#0A4C8C" }} /> Black Spot Heatmap</p>
            <p className={`text-xs mb-3 ${dark ? "text-slate-400" : "text-slate-500"}`}>Recurrence-coded GVP markers — click a marker for the full black-spot profile.</p>
            <HotspotMap dark={dark} onSelect={setSelected} selected={selected} compact />
          </div>
          <div className={glass}>
            <p className={`text-sm font-bold mb-2 ${dark ? "text-white" : "text-slate-900"}`}>Black Spot Profile</p>
            {selected ? (
              <div className="space-y-1.5 text-xs">
                {[
                  ["Black Spot ID", selected.id],
                  ["Ward", selected.ward],
                  ["Address / Landmark", selected.name],
                  ["GPS", `${selected.lat}, ${selected.lng}`],
                  ["Total Complaints", selected.complaints],
                  ["Last Cleaned", selected.lastCleaned],
                  ["AI Recurrence", `${selected.recurrence}%`],
                  ["Root Cause", selected.rootCause],
                  ["Nearby Market", nearbyOf(selected, "market")],
                  ["Nearby Dustbin", nearbyOf(selected, "bin")],
                  ["Assigned Worker", workerOf(selected)],
                  ["Current Status", gvpStatus(selected)],
                ].map(([k, v]) => (
                  <div key={k} className="flex justify-between gap-3 py-1 border-b" style={{ borderColor: dark ? "#334155" : "#eef2f6" }}>
                    <span style={{ color: dark ? "#94a3b8" : "#64748b" }}>{k}</span>
                    <span className={`font-semibold text-right ${dark ? "text-slate-100" : "text-slate-800"}`}>{v}</span>
                  </div>
                ))}
                <div className="pt-2"><RiskBadge risk={selected.risk} /></div>
              </div>
            ) : (
              <p className={`text-xs ${dark ? "text-slate-500" : "text-slate-400"}`}>Select a marker on the heatmap to view its black-spot profile.</p>
            )}
          </div>
        </div>

        {/* Complaint lifecycle timeline */}
        <div className={glass}>
          <div className="flex items-center justify-between flex-wrap gap-2 mb-3">
            <p className={`text-sm font-bold ${dark ? "text-white" : "text-slate-900"}`}>Complaint Lifecycle Timeline</p>
            <select value={tl ? tl.id : ""} onChange={e => setTimelineId(e.target.value)} className={selCls}>
              {fc.length === 0 && <option value="">No complaints</option>}
              {fc.slice(0, 40).map(c => <option key={c.id} value={c.id}>{c.id} — {c.status}</option>)}
            </select>
          </div>
          {tl ? (
            <ol className="relative">
              {TRACK_STAGES.map((st, i) => {
                const done = i < tlIdx, current = i === tlIdx, active = done || current;
                const at = tlHist[st];
                return (
                  <li key={st} className="flex gap-3 pb-4 last:pb-0">
                    <div className="flex flex-col items-center">
                      <div className="w-6 h-6 rounded-full flex items-center justify-center flex-shrink-0" style={{ background: active ? "#0B5D3B" : (dark ? "#334155" : "#e2e8f0") }}>
                        {done ? <Check size={13} color="#fff" /> : <span className="text-[10px] font-bold" style={{ color: active ? "#fff" : (dark ? "#94a3b8" : "#64748b") }}>{i + 1}</span>}
                      </div>
                      {i < TRACK_STAGES.length - 1 && <div className="w-0.5 flex-1 mt-1" style={{ background: done ? "#0B5D3B" : (dark ? "#334155" : "#e2e8f0"), minHeight: 18 }} />}
                    </div>
                    <div className="pb-1">
                      <p className={`text-sm font-semibold ${active ? (dark ? "text-white" : "text-slate-900") : (dark ? "text-slate-500" : "text-slate-400")}`}>{st}</p>
                      <p className="text-[11px]" style={{ color: dark ? "#94a3b8" : "#64748b" }}>{at ? fmt(at) : (current ? "In progress" : "Pending")}</p>
                    </div>
                  </li>
                );
              })}
            </ol>
          ) : (
            <p className={`text-xs ${dark ? "text-slate-500" : "text-slate-400"}`}>No complaints match the current filters.</p>
          )}
        </div>
      </div>
    </div>
  );
}

function HomePage({ dark, t, setView, selected, setSelected, complaints }) {
  return (
    <div>
      {/* HERO */}
      <div className="relative overflow-hidden" style={{ background: dark ? "linear-gradient(160deg,#0a1f16,#0a1830)" : "linear-gradient(160deg,#f0f7f2,#eaf2fb)" }}>
        <div className="max-w-4xl mx-auto px-4 sm:px-6 py-16 md:py-24">
          <div>
            <div className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-full text-xs font-semibold mb-5 ${dark ? "bg-emerald-900/40 text-emerald-300" : "bg-emerald-100 text-emerald-800"}`}>
              <Shield size={13} /> Greater Visakhapatnam Municipal Corporation
            </div>
            <h1 className={`text-4xl md:text-5xl font-extrabold leading-tight tracking-tight ${dark ? "text-white" : "text-slate-900"}`}>
              {t.heroTitle}
            </h1>
            <p className={`text-base mt-5 max-w-xl leading-relaxed ${dark ? "text-slate-300" : "text-slate-600"}`}>{t.heroSub}</p>
            <div className="flex flex-wrap gap-3 mt-8">
              <button onClick={() => setView("citizen")} className="flex items-center gap-2 px-5 py-3 rounded-xl text-white font-semibold text-sm shadow-lg shadow-emerald-900/10" style={{ background: "#0B5D3B" }}>
                <Camera size={16} /> {t.reportNow}
              </button>
              <button onClick={() => setView("ai")} className={`flex items-center gap-2 px-5 py-3 rounded-xl font-semibold text-sm border-2 ${dark ? "border-slate-700 text-white hover:bg-slate-800" : "border-slate-200 text-slate-700 hover:bg-white"}`}>
                <MapIcon size={16} /> {t.viewMap}
              </button>
            </div>
            <div className="flex items-center gap-6 mt-9">
              {[["8", "Zones Covered"], ["94%", "Prediction Accuracy"], ["24/7", "AI Monitoring"]].map(([n, l]) => (
                <div key={l}>
                  <p className={`text-2xl font-extrabold ${dark ? "text-white" : "text-slate-900"}`}>{n}</p>
                  <p className={`text-xs ${dark ? "text-slate-400" : "text-slate-500"}`}>{l}</p>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* LIVE STATS */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 -mt-8 relative z-10">
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          <StatCard icon={ClipboardList} label={t.totalComplaints} value="1,842" sub="+6.2% this month" accent="#0A4C8C" dark={dark} />
          <StatCard icon={AlertTriangle} label={t.activeGVPs} value="97" sub="Across 8 wards" accent="#E8A93B" dark={dark} />
          <StatCard icon={CheckCircle2} label={t.cleanedSpots} value="1,613" sub="87.6% resolution rate" accent="#0B5D3B" dark={dark} />
          <StatCard icon={ThermometerSun} label={t.highRisk} value="14" sub="Needs proactive action" accent="#D6432B" dark={dark} />
        </div>
      </div>

      {/* OVERVIEW */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 py-16">
        <SectionTitle eyebrow="How it works" title={t.overview} sub={t.overviewText} dark={dark} />
        <div className="grid md:grid-cols-4 gap-4">
          {[
            { icon: Camera, title: "1. Citizens Report", desc: "Photo, GPS, ward, garbage type & description submitted in seconds.", view: "citizen", cta: "Open Citizen Portal" },
            { icon: Brain, title: "2. AI Analyzes", desc: "ML models score recurrence risk using history, density, market proximity & rainfall.", view: "ai", cta: "See AI predictions" },
            { icon: MapIcon, title: "3. Hotspots Mapped", desc: "Clustering (K-Means/DBSCAN) reveals GVP clusters on a live GIS map.", view: "ai", cta: "View hotspot map" },
            { icon: Truck, title: "4. GVMC Acts", desc: "Officers dispatch crews, schedule cleaning, and prevent recurrence proactively.", view: "admin", cta: "Open Admin Dashboard" },
          ].map((s, i) => (
            <button
              key={i}
              type="button"
              onClick={() => setView(s.view)}
              aria-label={s.cta}
              className={`text-left w-full h-full rounded-2xl border p-5 shadow-sm transition-all cursor-pointer hover:shadow-md hover:-translate-y-0.5 focus:outline-none focus:ring-2 focus:ring-emerald-500 ${dark ? "bg-slate-800/60 border-slate-700 hover:border-emerald-700" : "bg-white border-slate-200 hover:border-emerald-300"}`}
            >
              <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-3" style={{ background: dark ? "#0f2a20" : "#eaf3ee" }}>
                <s.icon size={19} style={{ color: "#0B5D3B" }} />
              </div>
              <p className={`font-semibold text-sm ${dark ? "text-white" : "text-slate-900"}`}>{s.title}</p>
              <p className={`text-xs mt-1.5 leading-relaxed ${dark ? "text-slate-400" : "text-slate-500"}`}>{s.desc}</p>
              <span className="inline-flex items-center gap-1 text-xs font-semibold mt-3" style={{ color: "#0B5D3B" }}>
                {s.cta} <ChevronRight size={13} />
              </span>
            </button>
          ))}
        </div>
      </div>

      {/* MAP + TRENDS */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 pb-16 grid lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2">
          <SectionTitle eyebrow="Live GIS" title="Interactive Hotspot Map" sub="Click a marker to view AI-predicted recurrence, risk score and nearby facilities." dark={dark} />
          <HotspotMap dark={dark} onSelect={setSelected} selected={selected} />
          <ScheduleCrewSection selected={selected} onSelect={setSelected} dark={dark} />
        </div>
        <div>
          <SectionTitle eyebrow="Details" title="Selected Point" dark={dark} />
          {selected ? <HotspotDetailPanel hotspot={selected} dark={dark} onClose={() => setSelected(null)} /> : (
            <Card dark={dark} className="p-8 text-center">
              <MapPin size={28} className={`mx-auto mb-2 ${dark ? "text-slate-600" : "text-slate-300"}`} />
              <p className={`text-sm ${dark ? "text-slate-400" : "text-slate-500"}`}>Select a marker on the map to see full AI prediction details.</p>
            </Card>
          )}
        </div>
      </div>

      {/* AI Analytics Dashboard (replaces the old trend teaser) */}
      <AnalyticsDashboard complaints={complaints} dark={dark} selected={selected} setSelected={setSelected} />
    </div>
  );
}

/* ---------------------------------- CITIZEN PORTAL ---------------------------------- */
function CitizenPortal({ dark, t, user, complaints, addComplaint }) {
  const [tab, setTab] = useState("report");
  const [form, setForm] = useState({ ward: WARDS[0], type: GARBAGE_TYPES[0], address: "", desc: "", photo: null });
  const [submitted, setSubmitted] = useState(null);
  const [gps, setGps] = useState(null);
  const [showErrors, setShowErrors] = useState(false);
  const fileRef = useRef(null);

  const handleGPS = () => setGps({ lat: (17.68 + Math.random() * 0.15).toFixed(4), lng: (83.2 + Math.random() * 0.15).toFixed(4) });
  const handleFile = (e) => {
    const f = e.target.files?.[0];
    if (f) setForm({ ...form, photo: URL.createObjectURL(f) });
  };
  // Required fields — a complaint cannot be submitted until all are provided.
  const missing = {
    photo: !form.photo,
    gps: !gps,
    address: !form.address.trim(),
    desc: !form.desc.trim(),
  };
  const isIncomplete = missing.photo || missing.gps || missing.address || missing.desc;

  const handleSubmit = () => {
    if (isIncomplete) { setShowErrors(true); return; } // block until required info is filled
    const id = `GVMC-${2026000 + Math.floor(Math.random() * 900 + 100)}`;
    // Normalised complaint shape shared with the Officer dashboard.
    const complaint = {
      id,
      ward: form.ward,
      type: form.type,
      address: form.address || (form.ward.split(" - ")[1] || form.ward),
      description: form.desc,
      status: "Submitted",
      date: new Date().toISOString().slice(0, 10),
      citizen: user?.name || "Citizen",
      assignedTo: null,
      scheduledDate: null,
      statusHistory: [{ status: "Submitted", at: new Date().toISOString() }],
      gps,
      photo: form.photo,
      source: "citizen",
      createdAt: Date.now(),
    };
    addComplaint(complaint);       // -> instantly visible in the Officer Complaints Dashboard
    setSubmitted(complaint);
    setShowErrors(false);
    setForm({ ward: WARDS[0], type: GARBAGE_TYPES[0], address: "", desc: "", photo: null });
    setGps(null);
    setTab("track");
  };

  // The citizen sees only their own complaints (live from the shared store), newest first.
  const myComplaints = useMemo(
    () => complaints
      .filter(c => (user?.name ? c.citizen === user.name : true))
      .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0) || String(b.id).localeCompare(String(a.id))),
    [complaints, user?.name]
  );

  return (
    <div className="max-w-6xl mx-auto px-4 sm:px-6 py-10">
      <SectionTitle eyebrow="Citizen Portal" title="Report & Track Garbage Complaints" sub="Upload a photo with GPS location — GVMC's AI system verifies, prioritizes and schedules cleanup." dark={dark} />

      <div className={`flex gap-1 p-1 rounded-xl mb-6 w-fit ${dark ? "bg-slate-800" : "bg-slate-100"}`}>
        {[["report", "Report Garbage", Camera], ["track", "Track Complaint", Search], ["account", "My Account", Users]].map(([k, l, Icon]) => (
          <button key={k} onClick={() => setTab(k)} className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-sm font-semibold transition-colors ${tab === k ? "bg-white shadow text-emerald-700" : dark ? "text-slate-400" : "text-slate-500"}`}>
            <Icon size={14} /> {l}
          </button>
        ))}
      </div>

      {tab === "report" && (
        <div className="grid lg:grid-cols-3 gap-6">
          <Card dark={dark} className="p-6 lg:col-span-2">
            <p className={`text-sm font-bold mb-4 ${dark ? "text-white" : "text-slate-900"}`}>New Complaint Details</p>
            <div className="space-y-4">
              <div>
                <label className={`text-xs font-semibold ${dark ? "text-slate-400" : "text-slate-500"}`}>Photo Evidence <span style={{ color: "#D6432B" }}>*</span></label>
                <div onClick={() => fileRef.current?.click()} className="border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition-colors mt-1" style={{ borderColor: showErrors && missing.photo ? "#D6432B" : (dark ? "#334155" : "#e2e8f0") }}>
                  <input ref={fileRef} type="file" accept="image/*" className="hidden" onChange={handleFile} />
                  {form.photo ? (
                    <img src={form.photo} alt="Uploaded garbage" className="max-h-48 mx-auto rounded-lg object-cover" />
                  ) : (
                    <>
                      <Camera size={26} className={`mx-auto mb-2 ${dark ? "text-slate-500" : "text-slate-400"}`} />
                      <p className={`text-sm font-medium ${dark ? "text-slate-300" : "text-slate-600"}`}>Tap to upload photo evidence</p>
                      <p className={`text-xs mt-0.5 ${dark ? "text-slate-500" : "text-slate-400"}`}>JPG or PNG, up to 10MB</p>
                    </>
                  )}
                </div>
                {showErrors && missing.photo && <p className="text-xs mt-1 font-medium" style={{ color: "#D6432B" }}>Photo evidence is required.</p>}
              </div>

              <div>
                <label className={`text-xs font-semibold ${dark ? "text-slate-400" : "text-slate-500"}`}>GPS Location <span style={{ color: "#D6432B" }}>*</span></label>
                <button onClick={handleGPS} className="w-full flex items-center justify-between px-4 py-3 rounded-xl border text-sm font-medium mt-1" style={{ borderColor: showErrors && missing.gps ? "#D6432B" : (dark ? "#334155" : "#e2e8f0"), color: dark ? "#e2e8f0" : "#334155" }}>
                  <span className="flex items-center gap-2"><Navigation size={15} style={{ color: "#0A4C8C" }} /> {gps ? `GPS: ${gps.lat}, ${gps.lng}` : "Capture current GPS location"}</span>
                  {gps && <CheckCircle2 size={16} className="text-emerald-600" />}
                </button>
                {showErrors && missing.gps && <p className="text-xs mt-1 font-medium" style={{ color: "#D6432B" }}>Capture your GPS location to continue.</p>}
              </div>

              <div className="grid sm:grid-cols-2 gap-4">
                <div>
                  <label className={`text-xs font-semibold ${dark ? "text-slate-400" : "text-slate-500"}`}>Ward</label>
                  <select value={form.ward} onChange={e => setForm({ ...form, ward: e.target.value })} className={`w-full mt-1 px-3 py-2.5 rounded-xl border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white" : "bg-white border-slate-200 text-slate-900"}`}>
                    {WARDS.map(w => <option key={w}>{w}</option>)}
                  </select>
                </div>
                <div>
                  <label className={`text-xs font-semibold ${dark ? "text-slate-400" : "text-slate-500"}`}>Garbage Type</label>
                  <select value={form.type} onChange={e => setForm({ ...form, type: e.target.value })} className={`w-full mt-1 px-3 py-2.5 rounded-xl border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white" : "bg-white border-slate-200 text-slate-900"}`}>
                    {GARBAGE_TYPES.map(g => <option key={g}>{g}</option>)}
                  </select>
                </div>
              </div>

              <div>
                <label className={`text-xs font-semibold ${dark ? "text-slate-400" : "text-slate-500"}`}>Address / Landmark <span style={{ color: "#D6432B" }}>*</span></label>
                <input value={form.address} onChange={e => setForm({ ...form, address: e.target.value })} placeholder="e.g. Near RTC Complex, Dwaraka Nagar" className={`w-full mt-1 px-3 py-2.5 rounded-xl border text-sm ${dark ? "bg-slate-900 text-white placeholder:text-slate-500" : "bg-white text-slate-900 placeholder:text-slate-400"}`} style={{ borderColor: showErrors && missing.address ? "#D6432B" : (dark ? "#334155" : "#e2e8f0") }} />
                {showErrors && missing.address && <p className="text-xs mt-1 font-medium" style={{ color: "#D6432B" }}>Address or landmark is required.</p>}
              </div>

              <div>
                <label className={`text-xs font-semibold ${dark ? "text-slate-400" : "text-slate-500"}`}>Description <span style={{ color: "#D6432B" }}>*</span></label>
                <textarea value={form.desc} onChange={e => setForm({ ...form, desc: e.target.value })} rows={3} placeholder="Describe the issue..." className={`w-full mt-1 px-3 py-2.5 rounded-xl border text-sm resize-none ${dark ? "bg-slate-900 text-white placeholder:text-slate-500" : "bg-white text-slate-900 placeholder:text-slate-400"}`} style={{ borderColor: showErrors && missing.desc ? "#D6432B" : (dark ? "#334155" : "#e2e8f0") }} />
                {showErrors && missing.desc && <p className="text-xs mt-1 font-medium" style={{ color: "#D6432B" }}>Please describe the issue.</p>}
              </div>

              {showErrors && isIncomplete && (
                <div className="flex items-start gap-2 text-xs font-medium rounded-lg px-3 py-2" style={{ background: "#FCEBEB", color: "#791F1F" }}>
                  <AlertTriangle size={14} className="mt-0.5 flex-shrink-0" /> <span>Please complete all required fields (marked *) before submitting your complaint.</span>
                </div>
              )}

              <button
                onClick={handleSubmit}
                aria-disabled={isIncomplete}
                className="w-full py-3 rounded-xl text-white font-semibold text-sm flex items-center justify-center gap-2 transition-opacity"
                style={{ background: "#0B5D3B", opacity: isIncomplete ? 0.55 : 1, cursor: isIncomplete ? "not-allowed" : "pointer" }}
              >
                <Send size={15} /> Submit Complaint
              </button>
              <p className={`text-[11px] text-center ${dark ? "text-slate-500" : "text-slate-400"}`}>All fields marked <span style={{ color: "#D6432B" }}>*</span> are required.</p>
            </div>
          </Card>

          <div className="space-y-4">
            <Card dark={dark} className="p-5">
              <p className={`text-sm font-bold mb-2 ${dark ? "text-white" : "text-slate-900"}`}>What happens next?</p>
              {["Submitted", "Under Verification", "Cleaning Scheduled", "Cleaned"].map((s, i) => (
                <div key={s} className="flex items-center gap-3 py-2">
                  <div className="w-6 h-6 rounded-full flex items-center justify-center text-[11px] font-bold text-white flex-shrink-0" style={{ background: "#0B5D3B" }}>{i + 1}</div>
                  <p className={`text-sm ${dark ? "text-slate-300" : "text-slate-600"}`}>{s}</p>
                </div>
              ))}
            </Card>
            <Card dark={dark} className="p-5" >
              <p className={`text-sm font-bold mb-1 flex items-center gap-2 ${dark ? "text-white" : "text-slate-900"}`}><Bell size={14} style={{ color: "#0A4C8C" }} /> Notifications</p>
              <p className={`text-xs leading-relaxed ${dark ? "text-slate-400" : "text-slate-500"}`}>You'll receive SMS & email updates at every status change, including your unique Complaint ID.</p>
            </Card>
          </div>
        </div>
      )}

      {tab === "track" && (
        <div>
          {submitted && (
            <Card dark={dark} className="p-5 mb-5 border-l-4" style={{ borderLeftColor: "#0B5D3B" }}>
              <div className="flex items-center gap-2 mb-1">
                <CheckCircle2 size={16} className="text-emerald-600" />
                <p className={`text-sm font-bold ${dark ? "text-white" : "text-slate-900"}`}>Complaint submitted successfully!</p>
              </div>
              <p className={`text-xs ${dark ? "text-slate-400" : "text-slate-500"}`}>Your Complaint ID is <span className="font-mono font-bold">{submitted.id}</span> — save this for tracking.</p>
            </Card>
          )}

          <div className="flex items-center gap-2 mb-3">
            <span className="relative flex h-2.5 w-2.5">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full opacity-60" style={{ background: "#1D9E75" }}></span>
              <span className="relative inline-flex rounded-full h-2.5 w-2.5" style={{ background: "#1D9E75" }}></span>
            </span>
            <p className={`text-xs font-medium ${dark ? "text-slate-300" : "text-slate-600"}`}>Live tracking — updates automatically as GVMC officers act on your complaint.</p>
          </div>

          <div className="space-y-3">
            {myComplaints.length === 0 && (
              <Card dark={dark} className="p-8 text-center">
                <Search size={26} className={`mx-auto mb-2 ${dark ? "text-slate-600" : "text-slate-300"}`} />
                <p className={`text-sm ${dark ? "text-slate-400" : "text-slate-500"}`}>No complaints yet. Submit one from the Report Garbage tab to start tracking.</p>
              </Card>
            )}
            {myComplaints.map((c) => (
              <Card key={c.id} dark={dark} className="p-4">
                <div className="flex items-center justify-between flex-wrap gap-3">
                  <div>
                    <p className={`text-sm font-bold font-mono ${dark ? "text-white" : "text-slate-900"}`}>{c.id}</p>
                    <p className={`text-xs mt-0.5 ${dark ? "text-slate-400" : "text-slate-500"}`}>{c.ward} · {c.type} · {c.date}</p>
                  </div>
                  <StatusPill status={c.status} />
                </div>
                <ComplaintTracker complaint={c} dark={dark} />
                {c.assignedTo && (
                  <p className={`text-[11px] mt-3 flex items-center gap-1.5 ${dark ? "text-slate-400" : "text-slate-500"}`}>
                    <Truck size={12} style={{ color: "#0B5D3B" }} /> Crew: <span className="font-semibold">{c.assignedTo}</span>{c.scheduledDate ? ` · ${c.scheduledDate}` : ""}
                  </p>
                )}
              </Card>
            ))}
          </div>
        </div>
      )}

      {tab === "account" && (
        <Card dark={dark} className="p-6 max-w-md">
          <div className="flex items-center gap-3 mb-4">
            <div className="w-12 h-12 rounded-full flex items-center justify-center text-white font-bold" style={{ background: "#0A4C8C" }}>{(user?.name || "A. Sravani").split(" ").map(n => n[0]).join("").slice(0, 2)}</div>
            <div>
              <p className={`font-bold text-sm ${dark ? "text-white" : "text-slate-900"}`}>{user?.name || "A. Sravani"}</p>
              <p className={`text-xs ${dark ? "text-slate-400" : "text-slate-500"}`}>{user?.email || "sravani.a@example.com"}</p>
            </div>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div className={`rounded-xl p-3 text-center ${dark ? "bg-slate-900/60" : "bg-slate-50"}`}>
              <p className="text-xl font-bold" style={{ color: "#0A4C8C" }}>{myComplaints.length}</p>
              <p className={`text-xs ${dark ? "text-slate-400" : "text-slate-500"}`}>Total Reports</p>
            </div>
            <div className={`rounded-xl p-3 text-center ${dark ? "bg-slate-900/60" : "bg-slate-50"}`}>
              <p className="text-xl font-bold" style={{ color: "#0B5D3B" }}>{myComplaints.filter(c => ["Cleaned", "Closed"].includes(c.status)).length}</p>
              <p className={`text-xs ${dark ? "text-slate-400" : "text-slate-500"}`}>Resolved</p>
            </div>
          </div>
        </Card>
      )}
    </div>
  );
}

/* ---------------------------------- MANAGE COMPLAINT MODAL ---------------------------------- */
function ManageComplaintModal({ complaint, workers, dark, onClose, onSave }) {
  const [status, setStatus] = useState(complaint.status);
  const [assignedTo, setAssignedTo] = useState(complaint.assignedTo || "");
  const [scheduledDate, setScheduledDate] = useState(complaint.scheduledDate || "");

  const labelCls = `text-xs font-semibold ${dark ? "text-slate-400" : "text-slate-500"}`;
  const inputCls = `w-full mt-1 px-3 py-2.5 rounded-xl border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white" : "bg-white border-slate-200 text-slate-900"}`;

  const save = () => {
    // Assigning a worker or setting a schedule nudges the status forward if the
    // officer left it untouched, mirroring a realistic workflow.
    let nextStatus = status;
    if (scheduledDate && status === "Submitted") nextStatus = "Cleaning Scheduled";
    onSave({ ...complaint, status: nextStatus, assignedTo: assignedTo || null, scheduledDate: scheduledDate || null });
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(15,23,42,0.55)", zIndex: 2000 }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} className={`w-full max-w-md rounded-2xl p-6 relative ${dark ? "bg-slate-800" : "bg-white"}`}>
        <button onClick={onClose} className={`absolute top-4 right-4 p-1 rounded-lg ${dark ? "hover:bg-slate-700 text-slate-400" : "hover:bg-slate-100 text-slate-400"}`}><X size={16} /></button>
        <p className={`text-xs font-mono ${dark ? "text-slate-400" : "text-slate-500"}`}>{complaint.id}</p>
        <h3 className={`text-lg font-bold ${dark ? "text-white" : "text-slate-900"}`}>Manage Complaint</h3>
        <p className={`text-xs mb-4 ${dark ? "text-slate-400" : "text-slate-500"}`}>{complaint.ward} · {complaint.type}</p>

        <div className={`rounded-xl p-3 mb-4 text-xs ${dark ? "bg-slate-900/60 text-slate-300" : "bg-slate-50 text-slate-600"}`}>{complaint.address}</div>

        <div className="space-y-4">
          <div>
            <label className={labelCls}>Assign Sanitation Worker</label>
            <select value={assignedTo} onChange={e => setAssignedTo(e.target.value)} className={inputCls}>
              <option value="">Unassigned</option>
              {workers.map(w => <option key={w.id} value={w.name}>{w.name} — {w.ward}{w.status === "Off Duty" ? " (off duty)" : ""}</option>)}
            </select>
          </div>
          <div>
            <label className={labelCls}>Update Status</label>
            <select value={status} onChange={e => setStatus(e.target.value)} className={inputCls}>
              {STATUSES.map(s => <option key={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <label className={labelCls}>Schedule Cleaning Date</label>
            <input type="date" value={scheduledDate} onChange={e => setScheduledDate(e.target.value)} className={inputCls} />
          </div>
          <button onClick={save} className="w-full py-2.5 rounded-xl text-white text-sm font-semibold flex items-center justify-center gap-2" style={{ background: "#0B5D3B" }}>
            <CheckCircle2 size={15} /> Save Changes
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------- ADMIN DASHBOARD ---------------------------------- */
function AdminDashboard({ dark, complaints, setComplaints }) {
  const [tab, setTab] = useState("complaints");
  const [filterStatus, setFilterStatus] = useState("All");
  const [filterWard, setFilterWard] = useState("All");
  const [search, setSearch] = useState("");

  // Live, editable data — officers act on these and the UI reflects changes.
  // Complaints are the shared, app-wide store (props) — citizen submissions land here
  // automatically. Workers stay local to the dashboard.
  const [workers, setWorkers] = useState(WORKERS);
  const [managing, setManaging] = useState(null);
  const [toast, setToast] = useState("");
  const [newWorker, setNewWorker] = useState({ name: "", ward: WARDS[0].split(" - ")[0] });

  const flash = (msg) => { setToast(msg); setTimeout(() => setToast(""), 2600); };

  const filtered = complaints.filter(c =>
    (filterStatus === "All" || c.status === filterStatus) &&
    (filterWard === "All" || c.ward === filterWard) &&
    (c.id.toLowerCase().includes(search.toLowerCase()) || c.address.toLowerCase().includes(search.toLowerCase()))
  );

  // Live metrics
  const openCount = complaints.filter(c => !["Cleaned", "Closed"].includes(c.status)).length;
  const unassignedCount = complaints.filter(c => !c.assignedTo).length;
  const onDutyCount = workers.filter(w => w.status === "On Duty").length;
  const highRiskCount = HOTSPOTS.filter(h => h.risk === "High").length;

  // Live analytics derived from the current complaint set
  const wardCounts = WARDS.map(w => ({ ward: w.split(" - ")[1], complaints: complaints.filter(c => c.ward === w).length }));
  const statusCounts = STATUSES.map(s => ({ name: s, value: complaints.filter(c => c.status === s).length })).filter(x => x.value > 0);

  const saveComplaint = (updated) => {
    const old = complaints.find(c => c.id === updated.id);
    let next = updated;
    // Record a timestamped step whenever the status advances — this powers the
    // citizen's live tracker in the Citizen Portal.
    if (old && old.status !== updated.status) {
      const hist = Array.isArray(old.statusHistory) ? old.statusHistory : [{ status: old.status, at: old.date || new Date().toISOString() }];
      next = { ...updated, statusHistory: [...hist, { status: updated.status, at: new Date().toISOString() }] };
    }
    setComplaints(cs => cs.map(c => (c.id === next.id ? next : c)));
    flash(`Complaint ${updated.id} updated`);
  };

  const toggleDuty = (id) => {
    setWorkers(ws => ws.map(w => w.id === id ? { ...w, status: w.status === "On Duty" ? "Off Duty" : "On Duty" } : w));
  };

  const addWorker = () => {
    if (!newWorker.name.trim()) { flash("Enter a worker name first"); return; }
    const id = `W-${100 + workers.length + 1}`;
    setWorkers(ws => [...ws, { id, name: newWorker.name.trim(), ward: newWorker.ward, phone: "—", status: "On Duty", tasksToday: 0 }]);
    setNewWorker({ name: "", ward: WARDS[0].split(" - ")[0] });
    flash("Worker added");
  };

  const exportExcel = () => {
    try {
      const rows = complaints.map(c => ({
        "Complaint ID": c.id,
        Ward: c.ward,
        Type: c.type,
        Status: c.status,
        "Assigned To": c.assignedTo || "Unassigned",
        "Scheduled Cleaning": c.scheduledDate || "—",
        Address: c.address,
        Date: c.date,
        Citizen: c.citizen,
      }));
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Complaints");
      XLSX.writeFile(wb, "GVMC_Complaints_Report.xlsx");
      flash("Excel report downloaded");
    } catch (e) {
      flash("Excel export failed in this preview");
    }
  };

  const exportPDF = () => {
    const resolved = complaints.filter(c => ["Cleaned", "Closed"].includes(c.status)).length;
    const rowsHtml = complaints.map(c =>
      `<tr><td>${c.id}</td><td>${c.ward}</td><td>${c.type}</td><td>${c.status}</td><td>${c.assignedTo || "Unassigned"}</td><td>${c.scheduledDate || "—"}</td><td>${c.date}</td></tr>`
    ).join("");
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>GVMC Complaint Report</title>
      <style>
        body{font-family:Arial,Helvetica,sans-serif;padding:28px;color:#0f172a}
        h1{color:#0B5D3B;margin:0 0 4px}
        .meta{color:#64748b;font-size:12px;margin-bottom:16px}
        .cards{display:flex;gap:10px;margin-bottom:18px}
        .card{border:1px solid #e2e8f0;border-radius:10px;padding:10px 14px;font-size:12px}
        .card b{display:block;font-size:20px;color:#0A4C8C}
        table{width:100%;border-collapse:collapse;font-size:11px}
        th,td{border:1px solid #e2e8f0;padding:6px 8px;text-align:left}
        th{background:#f1f5f9}
      </style></head><body>
      <h1>Swachh GVMC — Complaint Report</h1>
      <div class="meta">Greater Visakhapatnam Municipal Corporation · Generated ${new Date().toLocaleString()}</div>
      <div class="cards">
        <div class="card"><b>${complaints.length}</b>Total complaints</div>
        <div class="card"><b>${openCount}</b>Open</div>
        <div class="card"><b>${resolved}</b>Resolved</div>
        <div class="card"><b>${unassignedCount}</b>Unassigned</div>
      </div>
      <table><thead><tr><th>ID</th><th>Ward</th><th>Type</th><th>Status</th><th>Assigned</th><th>Scheduled</th><th>Filed</th></tr></thead>
      <tbody>${rowsHtml}</tbody></table>
      </body></html>`;
    try {
      const iframe = document.createElement("iframe");
      Object.assign(iframe.style, { position: "fixed", right: "0", bottom: "0", width: "0", height: "0", border: "0" });
      document.body.appendChild(iframe);
      const doc = iframe.contentWindow.document;
      doc.open(); doc.write(html); doc.close();
      iframe.contentWindow.focus();
      setTimeout(() => {
        iframe.contentWindow.print();
        setTimeout(() => document.body.removeChild(iframe), 1200);
      }, 350);
      flash("Opening print dialog — choose 'Save as PDF'");
    } catch (e) {
      flash("PDF printing isn't available in this preview");
    }
  };

  const tabs = [
    ["complaints", "Complaints", ClipboardList],
    ["wards", "Wards & Workers", Truck],
    ["analytics", "Analytics", BarChart3],
    ["reports", "Reports", FileText],
  ];

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 py-10">
      <SectionTitle eyebrow="Admin Dashboard" title="GVMC Officer Control Center" sub="View all complaints, assign workers, update status, schedule cleaning, manage wards & workers, review analytics and download reports." dark={dark} />

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
        <StatCard icon={ClipboardList} label="Open Complaints" value={openCount} sub={`${unassignedCount} unassigned`} accent="#0A4C8C" dark={dark} />
        <StatCard icon={Truck} label="Active Workers" value={`${onDutyCount} / ${workers.length}`} sub="On duty now" accent="#0B5D3B" dark={dark} />
        <StatCard icon={Clock} label="Avg. Resolution" value="2.3 days" sub="-0.4d vs last month" accent="#E8A93B" dark={dark} />
        <StatCard icon={AlertTriangle} label="High-Risk GVPs" value={highRiskCount} sub="Predicted to recur" accent="#D6432B" dark={dark} />
      </div>

      <div className={`flex gap-1 p-1 rounded-xl mb-6 w-fit overflow-x-auto ${dark ? "bg-slate-800" : "bg-slate-100"}`}>
        {tabs.map(([k, l, Icon]) => (
          <button key={k} onClick={() => setTab(k)} className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-sm font-semibold whitespace-nowrap ${tab === k ? "bg-white shadow text-blue-700" : dark ? "text-slate-400" : "text-slate-500"}`}>
            <Icon size={14} /> {l}
          </button>
        ))}
      </div>

      {tab === "complaints" && (
        <div>
          <div className="flex flex-wrap gap-3 mb-4">
            <div className={`flex items-center gap-2 px-3 py-2 rounded-xl border flex-1 min-w-[200px] ${dark ? "bg-slate-800 border-slate-700" : "bg-white border-slate-200"}`}>
              <Search size={15} className={dark ? "text-slate-500" : "text-slate-400"} />
              <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by ID or address..." className={`bg-transparent text-sm outline-none w-full ${dark ? "text-white placeholder:text-slate-500" : "text-slate-900 placeholder:text-slate-400"}`} />
            </div>
            <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)} className={`px-3 py-2 rounded-xl border text-sm ${dark ? "bg-slate-800 border-slate-700 text-white" : "bg-white border-slate-200 text-slate-900"}`}>
              <option>All</option>{STATUSES.map(s => <option key={s}>{s}</option>)}
            </select>
            <select value={filterWard} onChange={e => setFilterWard(e.target.value)} className={`px-3 py-2 rounded-xl border text-sm ${dark ? "bg-slate-800 border-slate-700 text-white" : "bg-white border-slate-200 text-slate-900"}`}>
              <option>All</option>{WARDS.map(w => <option key={w}>{w}</option>)}
            </select>
          </div>

          <p className={`text-xs mb-2 ${dark ? "text-slate-400" : "text-slate-500"}`}>Showing {filtered.length} of {complaints.length} complaints</p>

          <Card dark={dark} className="overflow-hidden">
            <div className="overflow-x-auto max-h-[560px] overflow-y-auto">
              <table className="w-full text-sm">
                <thead className="sticky top-0">
                  <tr className={`text-left text-xs uppercase ${dark ? "text-slate-400 bg-slate-900" : "text-slate-500 bg-slate-100"}`}>
                    <th className="px-4 py-3 font-semibold">Complaint ID</th>
                    <th className="px-4 py-3 font-semibold">Ward</th>
                    <th className="px-4 py-3 font-semibold">Type</th>
                    <th className="px-4 py-3 font-semibold">Date</th>
                    <th className="px-4 py-3 font-semibold">Assigned</th>
                    <th className="px-4 py-3 font-semibold">Status</th>
                    <th className="px-4 py-3 font-semibold"></th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((c) => (
                    <tr key={c.id} className={`border-t ${dark ? "border-slate-700" : "border-slate-100"}`}>
                      <td className={`px-4 py-3 font-mono text-xs font-semibold ${dark ? "text-white" : "text-slate-900"}`}>{c.id}</td>
                      <td className={`px-4 py-3 text-xs ${dark ? "text-slate-300" : "text-slate-600"}`}>{c.ward.split(" - ")[1]}</td>
                      <td className={`px-4 py-3 text-xs ${dark ? "text-slate-300" : "text-slate-600"}`}>{c.type}</td>
                      <td className={`px-4 py-3 text-xs ${dark ? "text-slate-300" : "text-slate-600"}`}>{c.date}</td>
                      <td className={`px-4 py-3 text-xs ${dark ? "text-slate-300" : "text-slate-600"}`}>{c.assignedTo || <span className="italic opacity-60">Unassigned</span>}</td>
                      <td className="px-4 py-3"><StatusPill status={c.status} /></td>
                      <td className="px-4 py-3"><button onClick={() => setManaging(c)} className="text-xs font-semibold" style={{ color: "#0A4C8C" }}>Manage</button></td>
                    </tr>
                  ))}
                  {filtered.length === 0 && (
                    <tr><td colSpan={7} className={`px-4 py-8 text-center text-sm ${dark ? "text-slate-500" : "text-slate-400"}`}>No complaints match these filters.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </Card>
        </div>
      )}

      {tab === "wards" && (
        <div className="grid lg:grid-cols-2 gap-6">
          <Card dark={dark} className="p-5">
            <p className={`text-sm font-bold mb-4 ${dark ? "text-white" : "text-slate-900"}`}>Sanitation Workers</p>
            <div className="space-y-2">
              {workers.map(w => (
                <div key={w.id} className={`flex items-center justify-between p-3 rounded-xl ${dark ? "bg-slate-900/60" : "bg-slate-50"}`}>
                  <div className="flex items-center gap-3">
                    <div className="w-9 h-9 rounded-full flex items-center justify-center text-xs font-bold text-white" style={{ background: w.status === "On Duty" ? "#0B5D3B" : "#94a3b8" }}>{w.name.split(" ").map(n => n[0]).join("")}</div>
                    <div>
                      <p className={`text-sm font-semibold ${dark ? "text-white" : "text-slate-900"}`}>{w.name}</p>
                      <p className={`text-xs ${dark ? "text-slate-400" : "text-slate-500"}`}>{w.ward} · {w.tasksToday} tasks today</p>
                    </div>
                  </div>
                  <button onClick={() => toggleDuty(w.id)} className={`text-xs font-semibold px-2.5 py-1 rounded-full transition-colors ${w.status === "On Duty" ? "bg-emerald-50 text-emerald-700 hover:bg-emerald-100" : "bg-slate-200 text-slate-500 hover:bg-slate-300"}`}>{w.status}</button>
                </div>
              ))}
            </div>
            <div className={`mt-3 p-3 rounded-xl border-2 border-dashed ${dark ? "border-slate-700" : "border-slate-200"}`}>
              <p className={`text-xs font-semibold mb-2 ${dark ? "text-slate-300" : "text-slate-600"}`}>Add Worker</p>
              <div className="flex flex-col sm:flex-row gap-2">
                <input value={newWorker.name} onChange={e => setNewWorker({ ...newWorker, name: e.target.value })} placeholder="Worker name" className={`flex-1 px-3 py-2 rounded-lg border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white placeholder:text-slate-500" : "bg-white border-slate-200 text-slate-900 placeholder:text-slate-400"}`} />
                <select value={newWorker.ward} onChange={e => setNewWorker({ ...newWorker, ward: e.target.value })} className={`px-3 py-2 rounded-lg border text-sm ${dark ? "bg-slate-900 border-slate-700 text-white" : "bg-white border-slate-200 text-slate-900"}`}>
                  {WARDS.map(w => <option key={w} value={w.split(" - ")[0]}>{w.split(" - ")[0]}</option>)}
                </select>
                <button onClick={addWorker} className="px-3 py-2 rounded-lg text-white text-xs font-semibold flex items-center justify-center gap-1" style={{ background: "#0B5D3B" }}><Plus size={13} /> Add</button>
              </div>
            </div>
          </Card>
          <Card dark={dark} className="p-5">
            <p className={`text-sm font-bold mb-4 ${dark ? "text-white" : "text-slate-900"}`}>Ward Overview</p>
            <div className="space-y-2">
              {WARD_ANALYSIS.map(w => {
                const live = complaints.filter(c => c.ward.split(" - ")[1] === w.ward).length;
                return (
                  <div key={w.ward} className={`flex items-center justify-between p-3 rounded-xl ${dark ? "bg-slate-900/60" : "bg-slate-50"}`}>
                    <div>
                      <p className={`text-sm font-semibold ${dark ? "text-white" : "text-slate-900"}`}>{w.ward}</p>
                      <p className={`text-xs ${dark ? "text-slate-400" : "text-slate-500"}`}>{live} complaints on record</p>
                    </div>
                    <RiskBadge risk={w.risk > 70 ? "High" : w.risk > 40 ? "Medium" : "Low"} />
                  </div>
                );
              })}
            </div>
          </Card>
        </div>
      )}

      {tab === "analytics" && (
        <div className="grid lg:grid-cols-2 gap-6">
          <Card dark={dark} className="p-5">
            <p className={`text-sm font-bold mb-4 ${dark ? "text-white" : "text-slate-900"}`}>Complaints by Ward <span className={`font-normal ${dark ? "text-slate-500" : "text-slate-400"}`}>(live)</span></p>
            <div style={{ width: "100%", height: 260 }}>
              <ResponsiveContainer>
                <BarChart data={wardCounts} layout="vertical" margin={{ left: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#334155" : "#e2e8f0"} horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 11 }} axisLine={false} tickLine={false} allowDecimals={false} />
                  <YAxis type="category" dataKey="ward" width={110} tick={{ fontSize: 11, fill: dark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ borderRadius: 12, border: "none", fontSize: 12 }} />
                  <Bar dataKey="complaints" fill="#0A4C8C" radius={[0, 6, 6, 0]} barSize={16} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </Card>
          <Card dark={dark} className="p-5">
            <p className={`text-sm font-bold mb-4 ${dark ? "text-white" : "text-slate-900"}`}>Status Breakdown <span className={`font-normal ${dark ? "text-slate-500" : "text-slate-400"}`}>(live)</span></p>
            <div style={{ width: "100%", height: 260 }}>
              <ResponsiveContainer>
                <PieChart>
                  <Pie data={statusCounts} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={3}>
                    {statusCounts.map((e, i) => <Cell key={i} fill={(STATUS_COLORS[e.name] || {}).dot || PIE_COLORS[i % PIE_COLORS.length]} />)}
                  </Pie>
                  <Tooltip contentStyle={{ borderRadius: 12, border: "none", fontSize: 12 }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div className="flex flex-wrap gap-2 mt-2 justify-center">
              {statusCounts.map((e) => (
                <span key={e.name} className="flex items-center gap-1 text-xs" style={{ color: dark ? "#cbd5e1" : "#475569" }}>
                  <span className="w-2 h-2 rounded-full" style={{ background: (STATUS_COLORS[e.name] || {}).dot }}></span>{e.name} ({e.value})
                </span>
              ))}
            </div>
          </Card>
          <Card dark={dark} className="p-5 lg:col-span-2">
            <p className={`text-sm font-bold mb-4 ${dark ? "text-white" : "text-slate-900"}`}>Complaints vs Cleaned Trend</p>
            <div style={{ width: "100%", height: 240 }}>
              <ResponsiveContainer>
                <LineChart data={TREND_DATA}>
                  <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#334155" : "#e2e8f0"} vertical={false} />
                  <XAxis dataKey="month" tick={{ fontSize: 12 }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fontSize: 12 }} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ borderRadius: 12, border: "none", fontSize: 12 }} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Line type="monotone" dataKey="complaints" stroke="#D6432B" strokeWidth={2.5} dot={{ r: 3 }} name="Complaints" />
                  <Line type="monotone" dataKey="cleaned" stroke="#0B5D3B" strokeWidth={2.5} dot={{ r: 3 }} name="Cleaned" />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </Card>
        </div>
      )}

      {tab === "reports" && (
        <div className="grid sm:grid-cols-2 gap-6 max-w-2xl">
          <Card dark={dark} className="p-6 text-center">
            <FileText size={28} className="mx-auto mb-3" style={{ color: "#D6432B" }} />
            <p className={`text-sm font-bold ${dark ? "text-white" : "text-slate-900"}`}>Complaint PDF Report</p>
            <p className={`text-xs mt-1 mb-4 ${dark ? "text-slate-400" : "text-slate-500"}`}>Summary cards + full complaint table. Opens the print dialog so you can save as PDF.</p>
            <button onClick={exportPDF} className="w-full py-2.5 rounded-xl text-white text-sm font-semibold flex items-center justify-center gap-2" style={{ background: "#D6432B" }}><Download size={14} /> Download PDF</button>
          </Card>
          <Card dark={dark} className="p-6 text-center">
            <FileSpreadsheet size={28} className="mx-auto mb-3" style={{ color: "#0B5D3B" }} />
            <p className={`text-sm font-bold ${dark ? "text-white" : "text-slate-900"}`}>Raw Data Excel Export</p>
            <p className={`text-xs mt-1 mb-4 ${dark ? "text-slate-400" : "text-slate-500"}`}>Full complaint dataset with status, assignment & schedule as a real .xlsx file.</p>
            <button onClick={exportExcel} className="w-full py-2.5 rounded-xl text-white text-sm font-semibold flex items-center justify-center gap-2" style={{ background: "#0B5D3B" }}><Download size={14} /> Download Excel</button>
          </Card>
        </div>
      )}

      {managing && <ManageComplaintModal complaint={managing} workers={workers} dark={dark} onClose={() => setManaging(null)} onSave={saveComplaint} />}

      {toast && (
        <div className="fixed bottom-24 left-1/2 -translate-x-1/2 z-50 px-4 py-2.5 rounded-xl text-white text-sm font-semibold shadow-lg flex items-center gap-2" style={{ background: "#0B5D3B" }}>
          <CheckCircle2 size={15} /> {toast}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------- AI PREDICTION DASHBOARD ---------------------------------- */
function AIPredictionDashboard({ dark, selected, setSelected }) {
  const topRecurring = [...HOTSPOTS].sort((a, b) => b.recurrence - a.recurrence).slice(0, 5);
  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 py-10">
      <SectionTitle eyebrow="AI Engine" title="Garbage Recurrence Prediction Dashboard" sub="Random Forest / XGBoost recurrence scoring + K-Means / DBSCAN spatial clustering across historical complaints, cleaning logs, GPS, markets, dustbins, population density, rainfall and collection frequency." dark={dark} />

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
        <StatCard icon={Brain} label="Model Accuracy" value="94.2%" sub="XGBoost, validated" accent="#0A4C8C" dark={dark} />
        <StatCard icon={AlertTriangle} label="High-Risk Predicted" value="14 GVPs" sub=">70% recurrence" accent="#D6432B" dark={dark} />
        <StatCard icon={Layers} label="Clusters Found" value="6" sub="DBSCAN, eps=0.3km" accent="#E8A93B" dark={dark} />
        <StatCard icon={TrendingUp} label="Prevented Recurrence" value="31%" sub="vs. reactive model" accent="#0B5D3B" dark={dark} />
      </div>

      <div className="grid lg:grid-cols-3 gap-6 mb-8">
        <div className="lg:col-span-2">
          <Card dark={dark} className="p-5">
            <p className={`text-sm font-bold mb-3 flex items-center gap-2 ${dark ? "text-white" : "text-slate-900"}`}><MapIcon size={15} style={{ color: "#0A4C8C" }} /> Predicted Risk Heatmap</p>
            <HotspotMap dark={dark} onSelect={setSelected} selected={selected} />
          </Card>
        </div>
        <div>
          {selected ? <HotspotDetailPanel hotspot={selected} dark={dark} onClose={() => setSelected(null)} /> : (
            <Card dark={dark} className="p-5">
              <p className={`text-sm font-bold mb-3 ${dark ? "text-white" : "text-slate-900"}`}>Top Recurring Locations</p>
              <div className="space-y-2">
                {topRecurring.map((h, i) => (
                  <button key={h.id} onClick={() => setSelected(h)} className={`w-full text-left flex items-center justify-between p-2.5 rounded-xl ${dark ? "hover:bg-slate-700/40" : "hover:bg-slate-50"}`}>
                    <div className="flex items-center gap-2.5">
                      <span className={`w-6 h-6 rounded-full flex items-center justify-center text-[11px] font-bold ${dark ? "bg-slate-700 text-slate-300" : "bg-slate-100 text-slate-600"}`}>{i + 1}</span>
                      <div>
                        <p className={`text-xs font-semibold ${dark ? "text-white" : "text-slate-900"}`}>{h.name}</p>
                        <p className={`text-[11px] ${dark ? "text-slate-400" : "text-slate-500"}`}>{h.ward.split(" - ")[1]}</p>
                      </div>
                    </div>
                    <span className="text-xs font-bold" style={{ color: "#D6432B" }}>{h.recurrence}%</span>
                  </button>
                ))}
              </div>
            </Card>
          )}
        </div>
      </div>

      <div className="grid lg:grid-cols-2 gap-6 mb-8">
        <Card dark={dark} className="p-5">
          <p className={`text-sm font-bold mb-1 ${dark ? "text-white" : "text-slate-900"}`}>Feature Importance (Random Forest)</p>
          <p className={`text-xs mb-4 ${dark ? "text-slate-400" : "text-slate-500"}`}>Which signals drive the recurrence prediction most</p>
          <div style={{ width: "100%", height: 260 }}>
            <ResponsiveContainer>
              <BarChart data={FEATURE_IMPORTANCE} layout="vertical" margin={{ left: 10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={dark ? "#334155" : "#e2e8f0"} horizontal={false} />
                <XAxis type="number" domain={[0, 100]} tick={{ fontSize: 11 }} axisLine={false} tickLine={false} />
                <YAxis type="category" dataKey="feature" width={130} tick={{ fontSize: 11, fill: dark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ borderRadius: 12, border: "none", fontSize: 12 }} />
                <Bar dataKey="importance" fill="#0B5D3B" radius={[0, 6, 6, 0]} barSize={14} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </Card>
        <Card dark={dark} className="p-5">
          <p className={`text-sm font-bold mb-1 ${dark ? "text-white" : "text-slate-900"}`}>Ward-wise Risk Radar</p>
          <p className={`text-xs mb-4 ${dark ? "text-slate-400" : "text-slate-500"}`}>Predicted risk index (0–100) per ward</p>
          <div style={{ width: "100%", height: 260 }}>
            <ResponsiveContainer>
              <RadarChart data={WARD_ANALYSIS}>
                <PolarGrid stroke={dark ? "#334155" : "#e2e8f0"} />
                <PolarAngleAxis dataKey="ward" tick={{ fontSize: 9, fill: dark ? "#94a3b8" : "#64748b" }} />
                <PolarRadiusAxis tick={{ fontSize: 9 }} domain={[0, 100]} />
                <Radar dataKey="risk" stroke="#D6432B" fill="#D6432B" fillOpacity={0.25} strokeWidth={2} />
              </RadarChart>
            </ResponsiveContainer>
          </div>
        </Card>
      </div>

      <Card dark={dark} className="p-5">
        <p className={`text-sm font-bold mb-4 ${dark ? "text-white" : "text-slate-900"}`}>Recurrence Probability & Root Cause — All GVPs</p>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className={`text-left text-xs uppercase ${dark ? "text-slate-400 bg-slate-900/60" : "text-slate-500 bg-slate-50"}`}>
                <th className="px-3 py-2.5 font-semibold">Location</th>
                <th className="px-3 py-2.5 font-semibold">Ward</th>
                <th className="px-3 py-2.5 font-semibold">Recurrence</th>
                <th className="px-3 py-2.5 font-semibold">Risk</th>
                <th className="px-3 py-2.5 font-semibold">Root Cause (AI)</th>
              </tr>
            </thead>
            <tbody>
              {HOTSPOTS.map(h => (
                <tr key={h.id} className={`border-t cursor-pointer ${dark ? "border-slate-700 hover:bg-slate-700/30" : "border-slate-100 hover:bg-slate-50"}`} onClick={() => setSelected(h)}>
                  <td className={`px-3 py-2.5 font-semibold text-xs ${dark ? "text-white" : "text-slate-900"}`}>{h.name}</td>
                  <td className={`px-3 py-2.5 text-xs ${dark ? "text-slate-300" : "text-slate-600"}`}>{h.ward.split(" - ")[1]}</td>
                  <td className="px-3 py-2.5">
                    <div className="flex items-center gap-2">
                      <div className={`w-16 h-1.5 rounded-full ${dark ? "bg-slate-700" : "bg-slate-100"}`}>
                        <div className="h-1.5 rounded-full" style={{ width: `${h.recurrence}%`, background: h.recurrence > 70 ? "#D6432B" : h.recurrence > 40 ? "#E8A93B" : "#639922" }}></div>
                      </div>
                      <span className={`text-xs font-bold ${dark ? "text-white" : "text-slate-900"}`}>{h.recurrence}%</span>
                    </div>
                  </td>
                  <td className="px-3 py-2.5"><RiskBadge risk={h.risk} /></td>
                  <td className={`px-3 py-2.5 text-xs ${dark ? "text-slate-400" : "text-slate-500"}`}>{h.rootCause}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

/* ---------------------------------- CHATBOT ---------------------------------- */
function Chatbot({ dark }) {
  const [open, setOpen] = useState(false);
  const [msgs, setMsgs] = useState([{ from: "bot", text: "Namaste! I'm the GVMC Swachh Assistant. Ask me about reporting garbage, tracking a complaint, or your ward's cleanup schedule." }]);
  const [input, setInput] = useState("");

  const send = () => {
    if (!input.trim()) return;
    const userMsg = { from: "user", text: input };
    let reply = "You can report garbage from the Citizen Portal — upload a photo, allow GPS, and select your ward. You'll get a Complaint ID instantly.";
    if (/track|status/i.test(input)) reply = "To track a complaint, go to Citizen Portal → Track Complaint and enter your Complaint ID.";
    if (/risk|predict|ai/i.test(input)) reply = "Our AI Prediction Dashboard shows recurrence probability and risk scores for every Garbage Vulnerable Point using Random Forest & clustering models.";
    setMsgs(m => [...m, userMsg, { from: "bot", text: reply }]);
    setInput("");
  };

  return (
    <div className="fixed bottom-5 right-5 z-50">
      {open && (
        <div className={`w-80 mb-3 rounded-2xl border shadow-xl overflow-hidden ${dark ? "bg-slate-800 border-slate-700" : "bg-white border-slate-200"}`}>
          <div className="p-3.5 flex items-center gap-2 text-white" style={{ background: "#0B5D3B" }}>
            <MessageCircle size={16} /> <p className="text-sm font-semibold">Swachh Assistant</p>
          </div>
          <div className="h-64 overflow-y-auto p-3 space-y-2">
            {msgs.map((m, i) => (
              <div key={i} className={`max-w-[85%] text-xs px-3 py-2 rounded-xl ${m.from === "bot" ? (dark ? "bg-slate-700 text-slate-200" : "bg-slate-100 text-slate-700") : "ml-auto text-white"}`} style={m.from === "user" ? { background: "#0A4C8C" } : {}}>
                {m.text}
              </div>
            ))}
          </div>
          <div className={`flex items-center gap-2 p-2.5 border-t ${dark ? "border-slate-700" : "border-slate-200"}`}>
            <input value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => e.key === "Enter" && send()} placeholder="Type a message..." className={`flex-1 text-xs px-3 py-2 rounded-lg outline-none ${dark ? "bg-slate-900 text-white placeholder:text-slate-500" : "bg-slate-100 text-slate-900 placeholder:text-slate-400"}`} />
            <button onClick={send} className="p-2 rounded-lg text-white" style={{ background: "#0B5D3B" }}><Send size={13} /></button>
          </div>
        </div>
      )}
      <button onClick={() => setOpen(!open)} className="w-14 h-14 rounded-full shadow-lg flex items-center justify-center text-white" style={{ background: "linear-gradient(135deg,#0B5D3B,#0A4C8C)" }}>
        {open ? <X size={22} /> : <MessageCircle size={22} />}
      </button>
    </div>
  );
}

/* ---------------------------------- ROOT APP ---------------------------------- */
export default function GVPSystem() {
  const [view, setViewState] = useState("home");
  const [history, setHistory] = useState([]); // stack of previously-visited views

  // All navigation goes through this wrapper so we can offer a Back control.
  const setView = (next) => {
    setHistory(h => (next !== view ? [...h, view] : h));
    setViewState(next);
  };
  const goBack = () => {
    if (history.length === 0) return;
    const prev = history[history.length - 1];
    setHistory(history.slice(0, -1));
    setViewState(prev);
  };
  const VIEW_LABEL = { home: "Home", citizen: "Citizen Portal", admin: "Admin Dashboard", ai: "AI Predictions" };
  const [dark, setDark] = useState(false);
  const [lang, setLang] = useState("en");
  const [role, setRole] = useState("citizen");
  const [mobileOpen, setMobileOpen] = useState(false);
  const [selected, setSelected] = useState(null); // null → map opens on the full India overview
  const [accounts, setAccountsRaw] = useState(DEMO_ACCOUNTS);
  const [officers, setOfficersRaw] = useState(DEMO_OFFICERS);
  const [complaints, setComplaintsRaw] = useState(COMPLAINTS_SEED); // single app-wide source of truth
  const [user, setUser] = useState(null);
  const [authAudience, setAuthAudience] = useState(null); // null (closed) | 'citizen' | 'officer'
  const openAuth = (audience = "citizen") => setAuthAudience(audience);
  const closeAuth = () => setAuthAudience(null);
  const t = STR[lang];

  // On first mount: load registered citizens and officers, merge each with its demo
  // seed, then restore a previously signed-in session (citizen by email, officer by UID).
  useEffect(() => {
    let cancelled = false;
    (async () => {
      // Citizens (email-keyed)
      let mergedAcc = [...DEMO_ACCOUNTS];
      try {
        const res = await window.storage?.get("gvmc:accounts");
        if (res?.value) {
          JSON.parse(res.value).forEach((a) => {
            if (!mergedAcc.some((m) => m.email.toLowerCase() === a.email.toLowerCase())) mergedAcc.push(a);
          });
        }
      } catch (e) { /* demo citizens still work */ }

      // Officers (UID-keyed) — the separate officers database
      let mergedOff = [...DEMO_OFFICERS];
      try {
        const res = await window.storage?.get("gvmc:officers");
        if (res?.value) {
          JSON.parse(res.value).forEach((o) => {
            if (!mergedOff.some((m) => String(m.uid).toUpperCase() === String(o.uid).toUpperCase())) mergedOff.push(o);
          });
        }
      } catch (e) { /* demo officer still works */ }

      if (cancelled) return;
      setAccountsRaw(mergedAcc);
      setOfficersRaw(mergedOff);

      // Restore session (identifier only — never the password).
      try {
        const sess = await window.storage?.get("gvmc:session");
        if (!cancelled && sess?.value) {
          const s = JSON.parse(sess.value);
          let acct = null;
          if (s.role === "admin" && s.uid) acct = mergedOff.find((o) => String(o.uid).toUpperCase() === String(s.uid).toUpperCase());
          else if (s.email) acct = mergedAcc.find((a) => a.email.toLowerCase() === String(s.email).toLowerCase());
          if (acct) { setUser(acct); setRole(acct.role); }
        }
      } catch (e) { /* no active session */ }
    })();
    return () => { cancelled = true; };
  }, []);

  // Keeps state in sync AND persists only the user-created accounts
  // (demo accounts live in code, so there's no need to store them).
  const setAccounts = (next) => {
    setAccountsRaw((prev) => {
      const resolved = typeof next === "function" ? next(prev) : next;
      const custom = resolved.filter(
        (a) => !DEMO_ACCOUNTS.some((d) => d.email.toLowerCase() === a.email.toLowerCase())
      );
      try {
        window.storage?.set("gvmc:accounts", JSON.stringify(custom))?.catch(() => {});
      } catch (e) {
        // Persistence unavailable — accounts still work for this session.
      }
      return resolved;
    });
  };

  // Same idea for the officers database — persists only officers created at runtime.
  const setOfficers = (next) => {
    setOfficersRaw((prev) => {
      const resolved = typeof next === "function" ? next(prev) : next;
      const custom = resolved.filter(
        (o) => !DEMO_OFFICERS.some((d) => String(d.uid).toUpperCase() === String(o.uid).toUpperCase())
      );
      try {
        window.storage?.set("gvmc:officers", JSON.stringify(custom))?.catch(() => {});
      } catch (e) {
        // Persistence unavailable — officers still work for this session.
      }
      return resolved;
    });
  };

  // Shared complaints store. Citizen submissions and officer edits both flow through
  // here, so the Officer Complaints Dashboard always reflects the latest state.
  const setComplaints = (next) => {
    setComplaintsRaw((prev) => {
      const resolved = typeof next === "function" ? next(prev) : next;
      try {
        window.storage?.set("gvmc:complaints", JSON.stringify(resolved))?.catch(() => {});
      } catch (e) {
        // Persistence unavailable — complaints still work for this session.
      }
      return resolved;
    });
  };
  const addComplaint = (c) => setComplaints((prev) => [c, ...prev]);

  // Load any persisted complaints on first mount (survives reloads).
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await window.storage?.get("gvmc:complaints");
        if (!cancelled && res?.value) {
          const saved = JSON.parse(res.value);
          if (Array.isArray(saved) && saved.length) setComplaintsRaw(saved);
        }
      } catch (e) {
        // No saved complaints — start from the seed set.
      }
    })();
    return () => { cancelled = true; };
  }, []);

  const handleAuth = (account) => {
    setUser(account);
    setRole(account.role);
    closeAuth();
    setView(account.role === "admin" ? "admin" : "citizen");
    try {
      const sess = account.role === "admin"
        ? { role: "admin", uid: account.uid }
        : { role: "citizen", email: account.email };
      window.storage?.set("gvmc:session", JSON.stringify(sess))?.catch(() => {});
    } catch (e) {
      // Session persistence unavailable — still signed in for this session.
    }
  };
  const handleLogout = () => {
    setUser(null);
    setView("home");
    try {
      window.storage?.delete("gvmc:session")?.catch(() => {});
    } catch (e) {
      // Nothing to clear.
    }
  };

  return (
    <div className={dark ? "dark" : ""}>
      <div className={`min-h-screen font-sans transition-colors ${dark ? "bg-slate-900" : "bg-slate-50"}`}>
        <Navbar view={view} setView={setView} dark={dark} setDark={setDark} lang={lang} setLang={setLang} role={role} setRole={setRole} t={t} mobileOpen={mobileOpen} setMobileOpen={setMobileOpen} user={user} setShowAuth={() => openAuth("citizen")} onLogout={handleLogout} />
        {history.length > 0 && (
          <div className={`sticky top-16 z-30 border-b backdrop-blur-md ${dark ? "bg-slate-900/80 border-slate-800" : "bg-white/80 border-slate-200"}`}>
            <div className="max-w-7xl mx-auto px-4 sm:px-6 py-2 flex items-center gap-2">
              <button
                type="button"
                onClick={goBack}
                aria-label={`Back to ${VIEW_LABEL[history[history.length - 1]] || "previous"}`}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm font-semibold transition-colors ${dark ? "text-slate-200 hover:bg-slate-800" : "text-slate-700 hover:bg-slate-100"}`}
              >
                <ArrowLeft size={16} /> Back
              </button>
              <span className={`text-xs ${dark ? "text-slate-500" : "text-slate-400"}`}>
                to {VIEW_LABEL[history[history.length - 1]] || "previous page"}
              </span>
            </div>
          </div>
        )}
        {authAudience && <AuthModal audience={authAudience} dark={dark} onClose={closeAuth} onAuth={handleAuth} accounts={accounts} setAccounts={setAccounts} officers={officers} setOfficers={setOfficers} />}
        {view === "home" && <HomePage dark={dark} t={t} setView={setView} selected={selected} setSelected={setSelected} complaints={complaints} />}
        {view === "citizen" && (user ? <CitizenPortal dark={dark} t={t} user={user} complaints={complaints} addComplaint={addComplaint} /> : <AuthGate dark={dark} onLogin={() => openAuth("citizen")} label="the Citizen Portal" />)}
        {view === "admin" && (
          !user
            ? <AuthGate dark={dark} onLogin={() => openAuth("officer")} label="the Admin Dashboard (GVMC Officer login required)" />
            : user.role === "admin"
              ? <AdminDashboard dark={dark} complaints={complaints} setComplaints={setComplaints} />
              : <AccessDenied dark={dark} onSwitch={() => { handleLogout(); openAuth("officer"); }} />
        )}
        {view === "ai" && <AIPredictionDashboard dark={dark} selected={selected} setSelected={setSelected} />}

        <footer className={`border-t mt-10 py-8 text-center ${dark ? "border-slate-800 text-slate-500" : "border-slate-200 text-slate-500"}`}>
          <p className="text-xs">© 2026 Greater Visakhapatnam Municipal Corporation · Swachh GVMC AI-GVP System · Built with React, FastAPI, PostgreSQL & Scikit-learn</p>
          {(!user || user.role !== "admin") && (
            <button onClick={() => openAuth("officer")} className={`mt-3 inline-flex items-center gap-1.5 text-[11px] font-semibold ${dark ? "text-slate-400 hover:text-slate-200" : "text-slate-500 hover:text-slate-700"}`}>
              <Shield size={12} /> GVMC Officer Login
            </button>
          )}
        </footer>
        <Chatbot dark={dark} />
      </div>
    </div>
  );
}
