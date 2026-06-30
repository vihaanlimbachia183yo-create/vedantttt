import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import {
  Search, Plus, X, ChevronUp, ChevronDown, Info, Download, Upload, Save,
  RotateCcw, Printer, GripVertical, Filter, BarChart3, Settings, Trash2,
  AlertTriangle, ExternalLink, ArrowLeftRight, CheckCircle2, ChevronLeft,
  ChevronRight, User, ListChecks, Sparkles, Table2, FileSpreadsheet
} from 'lucide-react';
import * as XLSX from 'xlsx';

/* =========================================================================
   SAMPLE DATA LAYER
   -------------------------------------------------------------------------
   This entire block exists only so the app has something realistic to work
   with out of the box. EVERY cutoff number below is synthetically generated
   (seeded random, not scraped or official). Nothing here should be treated
   as an authoritative CET cutoff. The "Import Dataset" action in the
   Preferences tab is the supported path for loading real, official data
   (CSV/JSON) without touching this code -- see parseImportedDataset().
   ========================================================================= */

function mulberry32(seed) {
  let a = seed;
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
const round2 = (v) => Math.round(v * 100) / 100;

const COLLEGE_DEFS = [
  { capCode: '6011', name: 'College of Engineering, Pune', district: 'Pune', university: 'Savitribai Phule Pune University', autonomy: 'Autonomous', instType: 'Government', naac: 'A++', tier: 1, minority: 'None' },
  { capCode: '6012', name: 'Veermata Jijabai Technological Institute', district: 'Mumbai', university: 'Mumbai University', autonomy: 'Autonomous', instType: 'Government', naac: 'A++', tier: 2, minority: 'None' },
  { capCode: '6013', name: 'Institute of Chemical Technology', district: 'Mumbai', university: 'Mumbai University', autonomy: 'Autonomous', instType: 'Government', naac: 'A++', tier: 1, minority: 'None' },
  { capCode: '6014', name: 'Pune Institute of Computer Technology', district: 'Pune', university: 'Savitribai Phule Pune University', autonomy: 'Autonomous', instType: 'Private', naac: 'A+', tier: 4, minority: 'None' },
  { capCode: '6015', name: 'Vishwakarma Institute of Technology', district: 'Pune', university: 'Savitribai Phule Pune University', autonomy: 'Autonomous', instType: 'Private', naac: 'A+', tier: 5, minority: 'None' },
  { capCode: '6016', name: 'Sardar Patel College of Engineering', district: 'Mumbai', university: 'Mumbai University', autonomy: 'Non-Autonomous', instType: 'Government', naac: 'A', tier: 6, minority: 'None' },
  { capCode: '6017', name: 'Walchand College of Engineering', district: 'Sangli', university: 'Shivaji University, Kolhapur', autonomy: 'Autonomous', instType: 'Government', naac: 'A+', tier: 3, minority: 'None' },
  { capCode: '6018', name: 'Government College of Engineering, Chh. Sambhajinagar', district: 'Chh. Sambhajinagar', university: 'Dr. Babasaheb Ambedkar Marathwada University', autonomy: 'Non-Autonomous', instType: 'Government', naac: 'B++', tier: 9, minority: 'None' },
  { capCode: '6019', name: 'Government College of Engineering, Amravati', district: 'Amravati', university: 'Rashtrasant Tukadoji Maharaj Nagpur University', autonomy: 'Non-Autonomous', instType: 'Government', naac: 'B++', tier: 10, minority: 'None' },
  { capCode: '6020', name: 'Government College of Engineering, Karad', district: 'Satara', university: 'Shivaji University, Kolhapur', autonomy: 'Non-Autonomous', instType: 'Government', naac: 'B++', tier: 8, minority: 'None' },
  { capCode: '6021', name: 'Government College of Engineering, Jalgaon', district: 'Jalgaon', university: 'North Maharashtra University', autonomy: 'Non-Autonomous', instType: 'Government', naac: 'B+', tier: 12, minority: 'None' },
  { capCode: '6022', name: 'Cummins College of Engineering for Women', district: 'Pune', university: 'Savitribai Phule Pune University', autonomy: 'Autonomous', instType: 'Private', naac: 'A+', tier: 7, minority: 'None' },
  { capCode: '6023', name: "Anjuman-I-Islam Kalsekar Technical Campus", district: 'Raigad', university: 'Mumbai University', autonomy: 'Non-Autonomous', instType: 'Private', naac: 'B++', tier: 16, minority: 'Religious Minority' },
  { capCode: '6024', name: 'Father Agnel College of Engineering & Technology', district: 'Raigad', university: 'Mumbai University', autonomy: 'Non-Autonomous', instType: 'Private', naac: 'A', tier: 11, minority: 'Religious Minority' },
  { capCode: '6025', name: 'K. J. Somaiya College of Engineering', district: 'Mumbai', university: 'Mumbai University', autonomy: 'Autonomous', instType: 'Private', naac: 'A+', tier: 6, minority: 'None' },
  { capCode: '6026', name: 'Thadomal Shahani Engineering College', district: 'Mumbai', university: 'Mumbai University', autonomy: 'Non-Autonomous', instType: 'Private', naac: 'A', tier: 9, minority: 'None' },
  { capCode: '6027', name: 'Ramrao Adik Institute of Technology', district: 'Mumbai', university: 'Mumbai University', autonomy: 'Non-Autonomous', instType: 'Private', naac: 'B++', tier: 13, minority: 'None' },
  { capCode: '6028', name: 'Sinhgad College of Engineering', district: 'Pune', university: 'Savitribai Phule Pune University', autonomy: 'Non-Autonomous', instType: 'Private', naac: 'A', tier: 10, minority: 'None' },
  { capCode: '6029', name: 'MIT Academy of Engineering, Alandi', district: 'Pune', university: 'Savitribai Phule Pune University', autonomy: 'Autonomous', instType: 'Private', naac: 'A', tier: 8, minority: 'None' },
  { capCode: '6030', name: 'Bharati Vidyapeeth College of Engineering', district: 'Pune', university: 'Savitribai Phule Pune University', autonomy: 'Non-Autonomous', instType: 'Private', naac: 'A', tier: 11, minority: 'None' },
  { capCode: '6031', name: 'Shri Guru Gobind Singhji Inst. of Engg. & Technology', district: 'Nanded', university: 'Swami Ramanand Teerth Marathwada University', autonomy: 'Autonomous', instType: 'Government', naac: 'A+', tier: 4, minority: 'None' },
  { capCode: '6032', name: 'Datta Meghe College of Engineering', district: 'Thane', university: 'Mumbai University', autonomy: 'Non-Autonomous', instType: 'Private', naac: 'B++', tier: 15, minority: 'None' },
  { capCode: '6033', name: 'Army Institute of Technology', district: 'Pune', university: 'Savitribai Phule Pune University', autonomy: 'Autonomous', instType: 'Private', naac: 'A+', tier: 6, minority: 'None' },
  { capCode: '6034', name: 'Vidyalankar Institute of Technology', district: 'Mumbai', university: 'Mumbai University', autonomy: 'Non-Autonomous', instType: 'Private', naac: 'B++', tier: 14, minority: 'None' },
  { capCode: '6035', name: 'D. Y. Patil College of Engineering, Akurdi', district: 'Pune', university: 'Savitribai Phule Pune University', autonomy: 'Non-Autonomous', instType: 'Private', naac: 'A', tier: 12, minority: 'None' },
];

const BRANCHES = [
  { name: 'Computer Engineering', code: 'CO', demand: 1.00 },
  { name: 'Information Technology', code: 'IT', demand: 0.97 },
  { name: 'Artificial Intelligence & Data Science', code: 'AI', demand: 0.96 },
  { name: 'Electronics & Telecommunication Engg.', code: 'EN', demand: 0.85 },
  { name: 'Electrical Engineering', code: 'EE', demand: 0.80 },
  { name: 'Mechanical Engineering', code: 'ME', demand: 0.78 },
  { name: 'Civil Engineering', code: 'CE', demand: 0.68 },
  { name: 'Chemical Engineering', code: 'CH', demand: 0.72 },
];

const CATEGORIES = [
  { code: 'OPEN', adj: 0 },
  { code: 'EWS', adj: -1.5 },
  { code: 'OBC', adj: -4 },
  { code: 'SC', adj: -15 },
  { code: 'ST', adj: -22 },
];
const SEAT_TYPES = [
  { code: 'General', adj: 0 },
  { code: 'Ladies', adj: -1 },
];
const ALL_PROFILE_CATEGORIES = ['OPEN', 'EWS', 'OBC', 'SC', 'ST', 'VJ', 'NT1', 'NT2', 'NT3', 'SBC'];
const ALL_PROFILE_SEAT_TYPES = ['General', 'Ladies', 'TFWS', 'PWD', 'Defence', 'Orphan'];
const CANDIDATURE_TYPES = ['Maharashtra State Candidate', 'All India Candidate', 'J&K Migrant', 'CIWGC', 'Outside Maharashtra (OMS)'];
const MINORITY_OPTIONS = ['None', 'Linguistic Minority', 'Religious Minority'];

const rand = mulberry32(20260630);

function buildDataset() {
  const rows = [];
  let id = 1;
  COLLEGE_DEFS.forEach((college) => {
    BRANCHES.forEach((branch) => {
      const baseTier = clamp(99.9 - (college.tier - 1) * 1.4 - (1 - branch.demand) * 10, 30, 99.9);
      CATEGORIES.forEach((cat) => {
        SEAT_TYPES.forEach((seat) => {
          const r1 = round2(clamp(baseTier + cat.adj + seat.adj + (rand() - 0.5) * 1.2, 1, 99.99));
          const r2 = round2(clamp(r1 - rand() * 1.5, 1, 99.99));
          const r3 = round2(clamp(r2 - rand() * 1.5, 1, 99.99));
          const spot = rand() > 0.25 ? round2(clamp(r3 - rand() * 3, 1, 99.99)) : null;
          rows.push({
            id: id++,
            collegeName: college.name, capCode: college.capCode, district: college.district,
            university: college.university, autonomy: college.autonomy, instType: college.instType,
            naac: college.naac, minority: college.minority,
            branchName: branch.name, branchType: branch.code,
            category: cat.code, seatType: seat.code,
            cutoffs: { r1, r2, r3, spot },
          });
        });
      });
      const tr1 = round2(clamp(baseTier - 2.5 + (rand() - 0.5) * 1.2, 1, 99.99));
      const tr2 = round2(clamp(tr1 - rand() * 1.5, 1, 99.99));
      const tr3 = round2(clamp(tr2 - rand() * 1.5, 1, 99.99));
      const tspot = rand() > 0.25 ? round2(clamp(tr3 - rand() * 3, 1, 99.99)) : null;
      rows.push({
        id: id++,
        collegeName: college.name, capCode: college.capCode, district: college.district,
        university: college.university, autonomy: college.autonomy, instType: college.instType,
        naac: college.naac, minority: college.minority,
        branchName: branch.name, branchType: branch.code,
        category: 'OPEN', seatType: 'TFWS',
        cutoffs: { r1: tr1, r2: tr2, r3: tr3, spot: tspot },
      });
    });
  });
  return rows;
}

/** Maps an imported array of plain objects (parsed from CSV/JSON) into the
 *  app's internal row shape. Used by the Import Dataset action so a real
 *  official dataset can be dropped in without any code changes. Unknown /
 *  missing fields fall back sensibly so partial datasets still work. */
function parseImportedDataset(records) {
  return records.map((r, i) => ({
    id: 100000 + i,
    collegeName: String(r.collegeName || r.College || r['College Name'] || '').trim(),
    capCode: String(r.capCode || r.CAPCode || r['CAP Code'] || '').trim(),
    district: String(r.district || r.District || '').trim(),
    university: String(r.university || r.University || '').trim(),
    autonomy: String(r.autonomy || r.Autonomy || 'Non-Autonomous').trim(),
    instType: String(r.instType || r.Type || r['Govt/Private'] || 'Private').trim(),
    naac: r.naac || r.NAAC || null,
    minority: r.minority || r.Minority || 'None',
    branchName: String(r.branchName || r.Branch || r['Branch Name'] || '').trim(),
    branchType: String(r.branchType || r.BranchType || '').trim(),
    category: String(r.category || r.Category || 'OPEN').trim(),
    seatType: String(r.seatType || r.SeatType || r['Seat Type'] || 'General').trim(),
    cutoffs: {
      r1: r.r1 != null ? Number(r.r1) : (r.CAPRound1 != null ? Number(r.CAPRound1) : null),
      r2: r.r2 != null ? Number(r.r2) : (r.CAPRound2 != null ? Number(r.CAPRound2) : null),
      r3: r.r3 != null ? Number(r.r3) : (r.CAPRound3 != null ? Number(r.CAPRound3) : null),
      spot: r.spot != null ? Number(r.spot) : (r.Spot != null ? Number(r.Spot) : null),
    },
  })).filter(r => r.collegeName && r.branchName);
}

/* =========================================================================
   SMALL REUSABLE UI PRIMITIVES
   ========================================================================= */

const SAFETY_STYLES = {
  'Dream': 'bg-red-50 text-red-700 border-red-200',
  'Reach': 'bg-orange-50 text-orange-700 border-orange-200',
  'Competitive': 'bg-amber-50 text-amber-800 border-amber-200',
  'Safe': 'bg-green-50 text-green-700 border-green-200',
  'Very Safe': 'bg-teal-50 text-teal-700 border-teal-200',
};

function SafetyBadge({ label }) {
  if (!label) return <span className="text-xs text-slate-400">&mdash;</span>;
  return (
    <span className={`inline-block whitespace-nowrap rounded border px-2 py-0.5 text-xs font-semibold ${SAFETY_STYLES[label] || 'bg-slate-50 text-slate-600 border-slate-200'}`}>
      {label}
    </span>
  );
}

function Field({ label, children, hint }) {
  return (
    <label className="block">
      <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      {children}
      {hint && <span className="mt-1 block text-xs text-slate-400">{hint}</span>}
    </label>
  );
}

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-slate-900/50 p-4 sm:items-center">
      <div className={`w-full ${wide ? 'max-w-3xl' : 'max-w-lg'} rounded border border-slate-300 bg-white shadow-xl`}>
        <div className="flex items-center justify-between border-b border-slate-200 bg-blue-900 px-4 py-3">
          <h3 className="text-sm font-bold uppercase tracking-wide text-white">{title}</h3>
          <button onClick={onClose} className="rounded p-1 text-blue-100 hover:bg-blue-800 hover:text-white">
            <X size={18} />
          </button>
        </div>
        <div className="max-h-[75vh] overflow-y-auto p-4">{children}</div>
      </div>
    </div>
  );
}

function ConfirmModal({ message, onConfirm, onCancel }) {
  return (
    <Modal title="Please confirm" onClose={onCancel}>
      <div className="flex gap-3">
        <AlertTriangle className="mt-0.5 shrink-0 text-amber-500" size={22} />
        <p className="text-sm text-slate-700">{message}</p>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <button onClick={onCancel} className="rounded border border-slate-300 px-3 py-1.5 text-sm font-semibold text-slate-700 hover:bg-slate-50">Cancel</button>
        <button onClick={onConfirm} className="rounded bg-red-600 px-3 py-1.5 text-sm font-semibold text-white hover:bg-red-700">Confirm</button>
      </div>
    </Modal>
  );
}

function Toast({ toast }) {
  if (!toast) return null;
  const styles = {
    success: 'bg-green-700',
    warn: 'bg-amber-600',
    error: 'bg-red-700',
  };
  return (
    <div className={`fixed bottom-4 right-4 z-[60] flex items-center gap-2 rounded px-4 py-2.5 text-sm font-medium text-white shadow-lg ${styles[toast.type] || 'bg-slate-800'}`}>
      {toast.type === 'success' ? <CheckCircle2 size={16} /> : <AlertTriangle size={16} />}
      {toast.message}
    </div>
  );
}

function Pagination({ page, totalPages, onChange }) {
  if (totalPages <= 1) return null;
  return (
    <div className="flex items-center justify-between border-t border-slate-200 bg-slate-50 px-3 py-2 text-sm">
      <span className="text-slate-500">Page {page} of {totalPages}</span>
      <div className="flex gap-1">
        <button disabled={page <= 1} onClick={() => onChange(page - 1)} className="rounded border border-slate-300 bg-white p-1 disabled:opacity-40">
          <ChevronLeft size={16} />
        </button>
        <button disabled={page >= totalPages} onClick={() => onChange(page + 1)} className="rounded border border-slate-300 bg-white p-1 disabled:opacity-40">
          <ChevronRight size={16} />
        </button>
      </div>
    </div>
  );
}

/* =========================================================================
   MAIN APPLICATION
   ========================================================================= */

const DEFAULT_PROFILE = {
  name: '', percentile: '', category: 'OPEN', gender: 'Male',
  homeUniversity: 'Savitribai Phule Pune University', candidatureType: CANDIDATURE_TYPES[0],
  seatType: 'General', minorityStatus: 'None',
};
const DEFAULT_FILTERS = {
  district: '', university: '', instType: '', autonomy: '', branchType: '',
  minorityOnly: false, tfwsOnly: false, category: '', seatType: '', homeUniOnly: false,
};
const DEFAULT_MARGINS = { dream: 4, reach: 2, safe: -2, verySafe: -5 };
const PAGE_SIZE = 25;

export default function App() {
  const dataset = useMemo(() => buildDataset(), []);
  const [extraDataset, setExtraDataset] = useState([]); // rows from an imported file
  const fullDataset = useMemo(() => [...dataset, ...extraDataset], [dataset, extraDataset]);
  const rowById = useMemo(() => {
    const m = new Map();
    fullDataset.forEach(r => m.set(r.id, r));
    return m;
  }, [fullDataset]);

  const collegeMeta = useMemo(() => {
    const m = {};
    COLLEGE_DEFS.forEach((c, i) => {
      const lr = mulberry32(7000 + i * 13);
      m[c.name] = {
        avgPackageLPA: round2(3.5 + (25 - c.tier) * 0.35 + lr() * 1.5),
        feesPerYear: Math.round((60000 + c.tier * 4000 + lr() * 20000) / 1000) * 1000,
        intake: 60 + Math.round(lr() * 120),
      };
    });
    return m;
  }, []);

  const districts = useMemo(() => Array.from(new Set(fullDataset.map(r => r.district))).filter(Boolean).sort(), [fullDataset]);
  const universities = useMemo(() => Array.from(new Set(fullDataset.map(r => r.university))).filter(Boolean).sort(), [fullDataset]);
  const branchTypes = useMemo(() => Array.from(new Set(fullDataset.map(r => r.branchType))).filter(Boolean).sort(), [fullDataset]);

  const [profile, setProfile] = useState(DEFAULT_PROFILE);
  const [editingProfile, setEditingProfile] = useState(true);
  const [margins, setMargins] = useState(DEFAULT_MARGINS);
  const [refRound, setRefRound] = useState('r1');
  const [preferenceList, setPreferenceList] = useState([]); // [{optionRowId, notes}]
  const [activeTab, setActiveTab] = useState('dashboard');
  const [filters, setFilters] = useState(DEFAULT_FILTERS);
  const [searchText, setSearchText] = useState('');
  const [showAutocomplete, setShowAutocomplete] = useState(false);
  const [page, setPage] = useState(1);
  const [detailRowId, setDetailRowId] = useState(null);
  const [compareA, setCompareA] = useState(null);
  const [compareB, setCompareB] = useState(null);
  const [toast, setToast] = useState(null);
  const [showFilters, setShowFilters] = useState(false);
  const [saveStatus, setSaveStatus] = useState('idle');
  const [confirmAction, setConfirmAction] = useState(null);
  const dragIndexRef = useRef(null);
  const fileInputRef = useRef(null);
  const importInputRef = useRef(null);

  const showToast = useCallback((message, type = 'success') => {
    setToast({ message, type });
    const t = setTimeout(() => setToast(null), 2600);
    return () => clearTimeout(t);
  }, []);

  /* ---------- persistence via window.storage ---------- */
  const loadedRef = useRef(false);
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const p = await window.storage.get('cap-profile');
        if (p && !cancelled) setProfile(JSON.parse(p.value));
      } catch (e) { /* no saved profile yet */ }
      try {
        const pl = await window.storage.get('cap-preferences');
        if (pl && !cancelled) setPreferenceList(JSON.parse(pl.value));
      } catch (e) { /* no saved preferences yet */ }
      try {
        const m = await window.storage.get('cap-margins');
        if (m && !cancelled) setMargins(JSON.parse(m.value));
      } catch (e) { /* no saved margins yet */ }
      loadedRef.current = true;
    })();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!loadedRef.current) return;
    setSaveStatus('saving');
    const t = setTimeout(async () => {
      try {
        await window.storage.set('cap-profile', JSON.stringify(profile));
        await window.storage.set('cap-preferences', JSON.stringify(preferenceList));
        await window.storage.set('cap-margins', JSON.stringify(margins));
        setSaveStatus('saved');
      } catch (e) {
        setSaveStatus('error');
      }
    }, 700);
    return () => clearTimeout(t);
  }, [profile, preferenceList, margins]);

  useEffect(() => { setPage(1); }, [filters, searchText]);

  /* ---------- derived helpers ---------- */
  const percentileNum = profile.percentile === '' ? null : Number(profile.percentile);

  const diffFor = useCallback((row, round) => {
    if (!row || percentileNum == null || Number.isNaN(percentileNum)) return null;
    const c = row.cutoffs[round] != null ? row.cutoffs[round] : row.cutoffs.r1;
    if (c == null) return null;
    return round2(c - percentileNum);
  }, [percentileNum]);

  const safetyFor = useCallback((diff) => {
    if (diff == null) return null;
    if (diff >= margins.dream) return 'Dream';
    if (diff >= margins.reach) return 'Reach';
    if (diff >= margins.safe) return 'Competitive';
    if (diff >= margins.verySafe) return 'Safe';
    return 'Very Safe';
  }, [margins]);

  /* ---------- preference list actions ---------- */
  const addToPreferences = (rowId) => {
    if (preferenceList.some(p => p.optionRowId === rowId)) {
      showToast('That college + branch combination is already on your list.', 'warn');
      return;
    }
    setPreferenceList(prev => [...prev, { optionRowId: rowId, notes: '' }]);
    showToast('Added to your preference list.');
  };
  const removeFromPreferences = (rowId) => {
    setPreferenceList(prev => prev.filter(p => p.optionRowId !== rowId));
  };
  const updateNote = (rowId, text) => {
    setPreferenceList(prev => prev.map(p => p.optionRowId === rowId ? { ...p, notes: text } : p));
  };
  const moveItem = (index, dir) => {
    setPreferenceList(prev => {
      const next = [...prev];
      const j = index + dir;
      if (j < 0 || j >= next.length) return prev;
      [next[index], next[j]] = [next[j], next[index]];
      return next;
    });
  };
  const handleDrop = (index) => {
    const from = dragIndexRef.current;
    dragIndexRef.current = null;
    if (from == null || from === index) return;
    setPreferenceList(prev => {
      const next = [...prev];
      const [moved] = next.splice(from, 1);
      next.splice(index, 0, moved);
      return next;
    });
  };

  /* ---------- import / export ---------- */
  const downloadBlob = (content, filename, type) => {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const exportJSON = () => {
    const data = { profile, preferenceList, margins, refRound, exportedAt: new Date().toISOString() };
    downloadBlob(JSON.stringify(data, null, 2), 'cap-preference-list.json', 'application/json');
    showToast('Preference list exported as JSON.');
  };
  const importJSONFile = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        if (data.profile) setProfile(data.profile);
        if (Array.isArray(data.preferenceList)) setPreferenceList(data.preferenceList);
        if (data.margins) setMargins(data.margins);
        if (data.refRound) setRefRound(data.refRound);
        showToast('Preference list imported.');
      } catch (err) {
        showToast('That file is not a valid preference list export.', 'error');
      }
    };
    reader.readAsText(file);
  };
  const importDatasetFile = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        let records = [];
        if (file.name.toLowerCase().endsWith('.json')) {
          records = JSON.parse(e.target.result);
        } else {
          const wb = XLSX.read(e.target.result, { type: 'binary' });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          records = XLSX.utils.sheet_to_json(sheet);
        }
        const parsed = parseImportedDataset(records);
        setExtraDataset(parsed);
        showToast(`Imported ${parsed.length} rows into the college database.`);
      } catch (err) {
        showToast('Could not parse that dataset file.', 'error');
      }
    };
    if (file.name.toLowerCase().endsWith('.json')) reader.readAsText(file);
    else reader.readAsBinaryString(file);
  };

  const exportExcel = () => {
    const rows = preferenceList.map((p, i) => {
      const row = rowById.get(p.optionRowId);
      if (!row) return null;
      const d = diffFor(row, refRound);
      const s = safetyFor(d);
      return {
        'Preference': i + 1,
        'College': row.collegeName,
        'CAP Code': row.capCode,
        'Branch': row.branchName,
        'Category': row.category,
        'Seat Type': row.seatType,
        'Previous Year Cutoff': row.cutoffs[refRound] != null ? row.cutoffs[refRound] : '-',
        'Difference': d != null ? d : '-',
        'Notes': p.notes || '',
      };
    }).filter(Boolean);
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [{ wch: 10 }, { wch: 38 }, { wch: 9 }, { wch: 30 }, { wch: 9 }, { wch: 10 }, { wch: 14 }, { wch: 11 }, { wch: 30 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Preferences');
    XLSX.writeFile(wb, 'CAP_Preference_List.xlsx');
    showToast('Excel file downloaded.');
  };

  const resetAll = () => {
    setConfirmAction({
      message: 'This clears your entire preference list. This cannot be undone.',
      onConfirm: () => {
        setPreferenceList([]);
        setConfirmAction(null);
        showToast('Preference list cleared.', 'warn');
      },
    });
  };

  /* ---------- search / filter (Browse tab) ---------- */
  const filteredRows = useMemo(() => {
    const q = searchText.trim().toLowerCase();
    return fullDataset.filter(r => {
      if (filters.district && r.district !== filters.district) return false;
      if (filters.university && r.university !== filters.university) return false;
      if (filters.instType && r.instType !== filters.instType) return false;
      if (filters.autonomy && r.autonomy !== filters.autonomy) return false;
      if (filters.branchType && r.branchType !== filters.branchType) return false;
      if (filters.minorityOnly && r.minority === 'None') return false;
      if (filters.tfwsOnly && r.seatType !== 'TFWS') return false;
      if (filters.category && r.category !== filters.category) return false;
      if (filters.seatType && r.seatType !== filters.seatType) return false;
      if (filters.homeUniOnly && r.university !== profile.homeUniversity) return false;
      if (q) {
        const hay = `${r.collegeName} ${r.branchName} ${r.capCode} ${r.district}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [fullDataset, filters, searchText, profile.homeUniversity]);

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
  const pageRows = filteredRows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  const autocompleteOptions = useMemo(() => {
    const q = searchText.trim().toLowerCase();
    if (q.length < 1) return [];
    const set = new Set();
    for (const r of fullDataset) {
      if (set.size >= 8) break;
      if (r.collegeName.toLowerCase().includes(q)) set.add(r.collegeName);
    }
    for (const r of fullDataset) {
      if (set.size >= 8) break;
      if (r.branchName.toLowerCase().includes(q)) set.add(r.branchName);
    }
    return Array.from(set).slice(0, 8);
  }, [searchText, fullDataset]);

  /* ---------- suggestions (Dream/Reach/.../Very Safe) ---------- */
  const suggestionGroups = useMemo(() => {
    const groups = { 'Dream': [], 'Reach': [], 'Competitive': [], 'Safe': [], 'Very Safe': [] };
    if (percentileNum == null) return groups;
    const relevant = fullDataset.filter(r => r.category === profile.category && r.seatType === profile.seatType);
    relevant.forEach(r => {
      const d = diffFor(r, refRound);
      const s = safetyFor(d);
      if (s) groups[s].push({ row: r, diff: d });
    });
    Object.keys(groups).forEach(k => groups[k].sort((a, b) => Math.abs(a.diff) - Math.abs(b.diff)));
    return groups;
  }, [fullDataset, profile.category, profile.seatType, diffFor, safetyFor, percentileNum, refRound]);

  /* ---------- statistics ---------- */
  const stats = useMemo(() => {
    const rows = preferenceList.map(p => rowById.get(p.optionRowId)).filter(Boolean);
    const cutoffVals = rows.map(r => r.cutoffs[refRound]).filter(v => v != null);
    const counts = { 'Dream': 0, 'Reach': 0, 'Competitive': 0, 'Safe': 0, 'Very Safe': 0 };
    rows.forEach(r => {
      const s = safetyFor(diffFor(r, refRound));
      if (s) counts[s] += 1;
    });
    return {
      totalColleges: new Set(rows.map(r => r.collegeName)).size,
      totalBranches: new Set(rows.map(r => r.branchName)).size,
      highest: cutoffVals.length ? Math.max(...cutoffVals) : null,
      lowest: cutoffVals.length ? Math.min(...cutoffVals) : null,
      average: cutoffVals.length ? round2(cutoffVals.reduce((a, b) => a + b, 0) / cutoffVals.length) : null,
      counts,
      total: rows.length,
    };
  }, [preferenceList, rowById, refRound, safetyFor, diffFor]);

  /* =======================================================================
     RENDER
     ======================================================================= */
  const detailRow = detailRowId != null ? rowById.get(detailRowId) : null;
  const TABS = [
    { id: 'dashboard', label: 'My Profile', icon: User },
    { id: 'suggestions', label: 'Suggestions', icon: Sparkles },
    { id: 'browse', label: 'Browse & Filter', icon: Table2 },
    { id: 'preferences', label: 'Preference List', icon: ListChecks },
    { id: 'compare', label: 'Compare', icon: ArrowLeftRight },
    { id: 'stats', label: 'Statistics', icon: BarChart3 },
  ];

  return (
    <div className="min-h-full w-full bg-slate-100 font-sans text-slate-800">
      <style>{`
        @media print {
          .no-print { display: none !important; }
          .print-area { display: block !important; }
          body { background: white; }
        }
        .print-area { display: none; }
      `}</style>

      {/* Top tricolor accent strip */}
      <div className="flex h-1.5 w-full no-print">
        <div className="h-full flex-1 bg-orange-400" />
        <div className="h-full flex-1 bg-white border-y border-slate-200" />
        <div className="h-full flex-1 bg-green-600" />
      </div>

      {/* Header */}
      <header className="border-b border-slate-300 bg-blue-900 no-print">
        <div className="mx-auto max-w-7xl px-4 py-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div>
              <h1 className="text-lg font-bold tracking-tight text-white sm:text-xl">CAP Option Form Practice Tool</h1>
              <p className="text-xs text-blue-200">Maharashtra Engineering Admissions &mdash; unofficial practice workspace, not affiliated with DTE / State CET Cell</p>
            </div>
            <div className="flex items-center gap-2 text-xs text-blue-100">
              <span className={`inline-block h-2 w-2 rounded-full ${saveStatus === 'saving' ? 'bg-amber-400' : saveStatus === 'error' ? 'bg-red-400' : 'bg-green-400'}`} />
              {saveStatus === 'saving' ? 'Saving…' : saveStatus === 'error' ? 'Save failed' : 'Progress saved'}
            </div>
          </div>
        </div>
        <nav className="border-t border-blue-800 bg-blue-950">
          <div className="mx-auto flex max-w-7xl gap-1 overflow-x-auto px-2">
            {TABS.map(t => (
              <button
                key={t.id}
                onClick={() => setActiveTab(t.id)}
                className={`flex shrink-0 items-center gap-1.5 whitespace-nowrap border-b-2 px-3 py-2.5 text-sm font-semibold transition-colors ${activeTab === t.id ? 'border-orange-400 text-white' : 'border-transparent text-blue-200 hover:text-white'}`}
              >
                <t.icon size={15} /> {t.label}
                {t.id === 'preferences' && preferenceList.length > 0 && (
                  <span className="ml-1 rounded-full bg-orange-400 px-1.5 text-xs font-bold text-blue-950">{preferenceList.length}</span>
                )}
              </button>
            ))}
          </div>
        </nav>
      </header>

      <main className="mx-auto max-w-7xl px-3 py-4 no-print sm:px-4">
        {percentileNum == null && activeTab !== 'dashboard' && (
          <div className="mb-4 flex items-center gap-2 rounded border border-amber-300 bg-amber-50 px-3 py-2 text-sm text-amber-800">
            <Info size={16} className="shrink-0" /> Enter your MHT&nbsp;CET percentile on the My&nbsp;Profile tab to see cutoff comparisons and safety labels.
          </div>
        )}

        {activeTab === 'dashboard' && (
          <DashboardTab
            profile={profile} setProfile={setProfile}
            editingProfile={editingProfile} setEditingProfile={setEditingProfile}
            districts={districts} universities={universities}
            importDatasetFile={importDatasetFile} importInputRef={importInputRef}
            extraCount={extraDataset.length}
          />
        )}

        {activeTab === 'suggestions' && (
          <SuggestionsTab
            groups={suggestionGroups} margins={margins} setMargins={setMargins}
            refRound={refRound} setRefRound={setRefRound}
            onAdd={addToPreferences} onView={setDetailRowId}
            percentileSet={percentileNum != null}
            profile={profile}
          />
        )}

        {activeTab === 'browse' && (
          <BrowseTab
            filters={filters} setFilters={setFilters}
            searchText={searchText} setSearchText={setSearchText}
            showAutocomplete={showAutocomplete} setShowAutocomplete={setShowAutocomplete}
            autocompleteOptions={autocompleteOptions}
            districts={districts} universities={universities} branchTypes={branchTypes}
            pageRows={pageRows} page={page} setPage={setPage} totalPages={totalPages}
            totalCount={filteredRows.length}
            diffFor={diffFor} safetyFor={safetyFor} refRound={refRound} setRefRound={setRefRound}
            onAdd={addToPreferences} onView={setDetailRowId}
            showFilters={showFilters} setShowFilters={setShowFilters}
            preferenceIds={new Set(preferenceList.map(p => p.optionRowId))}
          />
        )}

        {activeTab === 'preferences' && (
          <PreferencesTab
            preferenceList={preferenceList} rowById={rowById}
            diffFor={diffFor} safetyFor={safetyFor} refRound={refRound}
            onRemove={removeFromPreferences} onNote={updateNote}
            onMove={moveItem} dragIndexRef={dragIndexRef} onDrop={handleDrop}
            onExportJSON={exportJSON} onExportExcel={exportExcel} onPrint={() => window.print()}
            onReset={resetAll} fileInputRef={fileInputRef} onImportJSON={importJSONFile}
            onView={setDetailRowId}
          />
        )}

        {activeTab === 'compare' && (
          <CompareTab
            dataset={fullDataset} compareA={compareA} compareB={compareB}
            setCompareA={setCompareA} setCompareB={setCompareB}
            collegeMeta={collegeMeta} refRound={refRound}
            diffFor={diffFor} safetyFor={safetyFor}
          />
        )}

        {activeTab === 'stats' && <StatsTab stats={stats} refRound={refRound} />}
      </main>

      {/* Printable preference list (only visible via @media print) */}
      <div className="print-area p-8">
        <h2 className="mb-1 text-xl font-bold">CAP Preference List &mdash; Practice Copy</h2>
        <p className="mb-4 text-sm text-slate-600">{profile.name || 'Candidate'} &middot; Percentile: {profile.percentile || '-'} &middot; Category: {profile.category} &middot; Seat Type: {profile.seatType}</p>
        <table className="w-full border-collapse text-sm">
          <thead>
            <tr className="border-b-2 border-black text-left">
              <th className="py-1 pr-2">#</th><th className="py-1 pr-2">College</th><th className="py-1 pr-2">Branch</th>
              <th className="py-1 pr-2">Cutoff</th><th className="py-1 pr-2">Diff</th><th className="py-1">Notes</th>
            </tr>
          </thead>
          <tbody>
            {preferenceList.map((p, i) => {
              const row = rowById.get(p.optionRowId);
              if (!row) return null;
              const d = diffFor(row, refRound);
              return (
                <tr key={p.optionRowId} className="border-b border-slate-300">
                  <td className="py-1 pr-2">{i + 1}</td>
                  <td className="py-1 pr-2">{row.collegeName} ({row.capCode})</td>
                  <td className="py-1 pr-2">{row.branchName}</td>
                  <td className="py-1 pr-2">{row.cutoffs[refRound] != null ? row.cutoffs[refRound] : '-'}</td>
                  <td className="py-1 pr-2">{d != null ? d : '-'}</td>
                  <td className="py-1">{p.notes}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <footer className="mx-auto max-w-7xl px-4 py-6 text-center text-xs text-slate-400 no-print">
        Independent practice tool. College list and cutoff figures are sample data for demonstration only &mdash; always verify against official DTE Maharashtra CAP publications before filling your real option form.
      </footer>

      {detailRow && (
        <DetailModal
          row={detailRow} meta={collegeMeta[detailRow.collegeName]}
          onClose={() => setDetailRowId(null)}
          onAdd={() => { addToPreferences(detailRow.id); }}
          onCompare={() => { setCompareA(detailRow.id); setActiveTab('compare'); setDetailRowId(null); }}
        />
      )}
      {confirmAction && (
        <ConfirmModal message={confirmAction.message} onConfirm={confirmAction.onConfirm} onCancel={() => setConfirmAction(null)} />
      )}
      <Toast toast={toast} />
    </div>
  );
}

/* =========================================================================
   TAB: DASHBOARD
   ========================================================================= */
function DashboardTab({ profile, setProfile, editingProfile, setEditingProfile, districts, universities, importDatasetFile, importInputRef, extraCount }) {
  const set = (k, v) => setProfile(prev => ({ ...prev, [k]: v }));
  return (
    <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
      <div className="rounded border border-slate-300 bg-white lg:col-span-2">
        <div className="flex items-center justify-between border-b border-slate-200 bg-slate-50 px-4 py-2.5">
          <h2 className="text-sm font-bold uppercase tracking-wide text-blue-900">Candidate Profile</h2>
          <button onClick={() => setEditingProfile(e => !e)} className="rounded border border-blue-700 px-2.5 py-1 text-xs font-semibold text-blue-700 hover:bg-blue-50">
            {editingProfile ? 'Done Editing' : 'Edit Profile'}
          </button>
        </div>
        <div className="p-4">
          {editingProfile ? (
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              <Field label="Full Name">
                <input value={profile.name} onChange={e => set('name', e.target.value)} placeholder="As per CET hall ticket" className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-sm focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600" />
              </Field>
              <Field label="MHT CET Percentile" hint="0–100, two decimals">
                <input type="number" min="0" max="100" step="0.01" value={profile.percentile} onChange={e => set('percentile', e.target.value)} placeholder="e.g. 92.45" className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-sm focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600" />
              </Field>
              <Field label="Category">
                <select value={profile.category} onChange={e => set('category', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-sm focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600">
                  {ALL_PROFILE_CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </Field>
              <Field label="Gender">
                <select value={profile.gender} onChange={e => set('gender', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-sm focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600">
                  <option>Male</option><option>Female</option><option>Other</option>
                </select>
              </Field>
              <Field label="Home University">
                <select value={profile.homeUniversity} onChange={e => set('homeUniversity', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-sm focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600">
                  {universities.map(u => <option key={u} value={u}>{u}</option>)}
                </select>
              </Field>
              <Field label="Candidature Type">
                <select value={profile.candidatureType} onChange={e => set('candidatureType', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-sm focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600">
                  {CANDIDATURE_TYPES.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </Field>
              <Field label="Seat Type">
                <select value={profile.seatType} onChange={e => set('seatType', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-sm focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600">
                  {ALL_PROFILE_SEAT_TYPES.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </Field>
              <Field label="Minority Status">
                <select value={profile.minorityStatus} onChange={e => set('minorityStatus', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-sm focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600">
                  {MINORITY_OPTIONS.map(m => <option key={m} value={m}>{m}</option>)}
                </select>
              </Field>
            </div>
          ) : (
            <dl className="grid grid-cols-1 gap-x-4 gap-y-3 sm:grid-cols-2">
              {[
                ['Name', profile.name || '—'], ['Percentile', profile.percentile || '—'],
                ['Category', profile.category], ['Gender', profile.gender],
                ['Home University', profile.homeUniversity], ['Candidature Type', profile.candidatureType],
                ['Seat Type', profile.seatType], ['Minority Status', profile.minorityStatus],
              ].map(([k, v]) => (
                <div key={k} className="border-b border-slate-100 pb-1.5">
                  <dt className="text-xs font-semibold uppercase tracking-wide text-slate-500">{k}</dt>
                  <dd className="text-sm text-slate-800">{v}</dd>
                </div>
              ))}
            </dl>
          )}
        </div>
      </div>

      <div className="space-y-4">
        <div className="rounded border border-slate-300 bg-white">
          <div className="border-b border-slate-200 bg-slate-50 px-4 py-2.5">
            <h2 className="text-sm font-bold uppercase tracking-wide text-blue-900">Import Official Dataset</h2>
          </div>
          <div className="space-y-2 p-4 text-sm text-slate-600">
            <p>Drop in a real CET CAP cutoff dataset (CSV/XLSX/JSON) and it merges straight into the college database &mdash; no code changes needed.</p>
            <p className="text-xs text-slate-500">Expected columns: College, CAP Code, District, University, Branch, Category, Seat Type, CAPRound1, CAPRound2, CAPRound3, Spot.</p>
            <input ref={importInputRef} type="file" accept=".csv,.xlsx,.xls,.json" className="hidden" onChange={e => { if (e.target.files[0]) importDatasetFile(e.target.files[0]); e.target.value = ''; }} />
            <button onClick={() => importInputRef.current && importInputRef.current.click()} className="flex items-center gap-1.5 rounded bg-blue-700 px-3 py-1.5 text-xs font-semibold text-white hover:bg-blue-800">
              <Upload size={14} /> Import Dataset File
            </button>
            {extraCount > 0 && <p className="text-xs font-semibold text-green-700">{extraCount} imported rows currently active.</p>}
          </div>
        </div>
        <div className="rounded border border-slate-300 bg-white p-4 text-xs text-slate-500">
          <p className="font-semibold text-slate-600">About this tool</p>
          <p className="mt-1">All college names, districts and cutoff numbers shown by default are illustrative sample data for practicing the option-form workflow. They are not official DTE Maharashtra figures.</p>
        </div>
      </div>
    </div>
  );
}

/* =========================================================================
   TAB: SUGGESTIONS
   ========================================================================= */
function SuggestionsTab({ groups, margins, setMargins, refRound, setRefRound, onAdd, onView, percentileSet, profile }) {
  const order = ['Dream', 'Reach', 'Competitive', 'Safe', 'Very Safe'];
  return (
    <div className="space-y-4">
      <div className="rounded border border-slate-300 bg-white p-4">
        <div className="flex flex-wrap items-end gap-4">
          <div>
            <h2 className="text-sm font-bold uppercase tracking-wide text-blue-900">Prediction Settings</h2>
            <p className="text-xs text-slate-500">Tune how far above/below your percentile a cutoff counts as each label (percentile points).</p>
          </div>
          <Field label="Compare against round">
            <select value={refRound} onChange={e => setRefRound(e.target.value)} className="rounded border border-slate-300 bg-white px-2 py-1 text-sm">
              <option value="r1">CAP Round I</option><option value="r2">CAP Round II</option>
              <option value="r3">CAP Round III</option><option value="spot">Spot Round</option>
            </select>
          </Field>
          {[['dream', 'Dream ≥'], ['reach', 'Reach ≥'], ['safe', 'Safe ≥'], ['verySafe', 'Very Safe <']].map(([k, lbl]) => (
            <Field key={k} label={lbl}>
              <input type="number" step="0.5" value={margins[k]} onChange={e => setMargins(m => ({ ...m, [k]: Number(e.target.value) }))} className="w-20 rounded border border-slate-300 px-2 py-1 text-sm" />
            </Field>
          ))}
        </div>
        <p className="mt-3 text-xs text-slate-500">Showing colleges for category <strong>{profile.category}</strong>, seat type <strong>{profile.seatType}</strong> from your profile.</p>
      </div>

      {!percentileSet ? (
        <div className="rounded border border-amber-300 bg-amber-50 p-6 text-center text-sm text-amber-800">
          Enter your percentile on the My Profile tab to generate suggestions.
        </div>
      ) : (
        order.map(label => (
          <div key={label} className="rounded border border-slate-300 bg-white">
            <div className={`flex items-center justify-between border-b border-slate-200 px-4 py-2 ${SAFETY_STYLES[label]}`}>
              <h3 className="text-sm font-bold">{label} <span className="font-normal">({groups[label].length} matches)</span></h3>
            </div>
            {groups[label].length === 0 ? (
              <p className="p-4 text-sm text-slate-400">No sample colleges fall in this band for your category/seat type.</p>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full min-w-[640px] text-sm">
                  <tbody>
                    {groups[label].slice(0, 8).map(({ row, diff }) => (
                      <tr key={row.id} className="border-b border-slate-100 last:border-0 hover:bg-slate-50">
                        <td className="px-3 py-2">
                          <div className="font-semibold text-slate-800">{row.collegeName}</div>
                          <div className="text-xs text-slate-500">{row.branchName} &middot; {row.district}</div>
                        </td>
                        <td className="px-3 py-2 text-right tabular-nums">{row.cutoffs[refRound] != null ? row.cutoffs[refRound] : '-'}</td>
                        <td className="px-3 py-2 text-right tabular-nums text-slate-500">{diff != null ? (diff > 0 ? '+' : '') + diff : '-'}</td>
                        <td className="px-3 py-2 text-right">
                          <button onClick={() => onView(row.id)} className="mr-2 text-xs font-semibold text-blue-700 hover:underline">View</button>
                          <button onClick={() => onAdd(row.id)} className="inline-flex items-center gap-1 rounded bg-blue-700 px-2 py-1 text-xs font-semibold text-white hover:bg-blue-800"><Plus size={12} /> Add</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        ))
      )}
    </div>
  );
}

/* =========================================================================
   TAB: BROWSE & FILTER
   ========================================================================= */
function BrowseTab(props) {
  const {
    filters, setFilters, searchText, setSearchText, showAutocomplete, setShowAutocomplete,
    autocompleteOptions, districts, universities, branchTypes, pageRows, page, setPage,
    totalPages, totalCount, diffFor, safetyFor, refRound, setRefRound, onAdd, onView,
    showFilters, setShowFilters, preferenceIds,
  } = props;
  const setF = (k, v) => setFilters(prev => ({ ...prev, [k]: v }));

  return (
    <div className="flex flex-col gap-4 lg:flex-row">
      <aside className="lg:w-64 lg:shrink-0">
        <button onClick={() => setShowFilters(s => !s)} className="mb-2 flex w-full items-center justify-center gap-1.5 rounded border border-slate-300 bg-white py-2 text-sm font-semibold text-slate-700 lg:hidden">
          <Filter size={15} /> {showFilters ? 'Hide Filters' : 'Show Filters'}
        </button>
        <div className={`${showFilters ? 'block' : 'hidden'} space-y-3 rounded border border-slate-300 bg-white p-3 lg:block`}>
          <h3 className="text-xs font-bold uppercase tracking-wide text-blue-900">Filters</h3>
          <Field label="District">
            <select value={filters.district} onChange={e => setF('district', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2 py-1.5 text-sm">
              <option value="">All Districts</option>{districts.map(d => <option key={d} value={d}>{d}</option>)}
            </select>
          </Field>
          <Field label="University">
            <select value={filters.university} onChange={e => setF('university', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2 py-1.5 text-sm">
              <option value="">All Universities</option>{universities.map(u => <option key={u} value={u}>{u}</option>)}
            </select>
          </Field>
          <Field label="Branch">
            <select value={filters.branchType} onChange={e => setF('branchType', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2 py-1.5 text-sm">
              <option value="">All Branches</option>{branchTypes.map(b => <option key={b} value={b}>{BRANCHES_LOOKUP[b] || b}</option>)}
            </select>
          </Field>
          <Field label="Government / Private">
            <select value={filters.instType} onChange={e => setF('instType', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2 py-1.5 text-sm">
              <option value="">Any</option><option>Government</option><option>Private</option><option>Aided</option>
            </select>
          </Field>
          <Field label="Autonomous">
            <select value={filters.autonomy} onChange={e => setF('autonomy', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2 py-1.5 text-sm">
              <option value="">Any</option><option>Autonomous</option><option>Non-Autonomous</option>
            </select>
          </Field>
          <Field label="Category">
            <select value={filters.category} onChange={e => setF('category', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2 py-1.5 text-sm">
              <option value="">Any</option>{CATEGORIES.map(c => <option key={c.code} value={c.code}>{c.code}</option>)}
            </select>
          </Field>
          <Field label="Seat Type">
            <select value={filters.seatType} onChange={e => setF('seatType', e.target.value)} className="w-full rounded border border-slate-300 bg-white px-2 py-1.5 text-sm">
              <option value="">Any</option><option>General</option><option>Ladies</option><option>TFWS</option>
            </select>
          </Field>
          <label className="flex items-center gap-2 text-sm text-slate-700">
            <input type="checkbox" checked={filters.minorityOnly} onChange={e => setF('minorityOnly', e.target.checked)} /> Minority institutes only
          </label>
          <label className="flex items-center gap-2 text-sm text-slate-700">
            <input type="checkbox" checked={filters.tfwsOnly} onChange={e => setF('tfwsOnly', e.target.checked)} /> TFWS seats only
          </label>
          <label className="flex items-center gap-2 text-sm text-slate-700">
            <input type="checkbox" checked={filters.homeUniOnly} onChange={e => setF('homeUniOnly', e.target.checked)} /> My home university only
          </label>
          <button onClick={() => setFilters(DEFAULT_FILTERS)} className="w-full rounded border border-slate-300 py-1.5 text-xs font-semibold text-slate-600 hover:bg-slate-50">Clear all filters</button>
        </div>
      </aside>

      <div className="min-w-0 flex-1">
        <div className="mb-3 flex flex-wrap items-center gap-3">
          <div className="relative flex-1 min-w-[220px]">
            <Search size={15} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              value={searchText}
              onChange={e => { setSearchText(e.target.value); setShowAutocomplete(true); }}
              onFocus={() => setShowAutocomplete(true)}
              onBlur={() => setTimeout(() => setShowAutocomplete(false), 150)}
              placeholder="Search college, branch, CAP code or district…"
              className="w-full rounded border border-slate-300 py-2 pl-8 pr-3 text-sm focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600"
            />
            {showAutocomplete && autocompleteOptions.length > 0 && (
              <div className="absolute z-10 mt-1 w-full rounded border border-slate-300 bg-white shadow-lg">
                {autocompleteOptions.map(o => (
                  <button key={o} onMouseDown={() => { setSearchText(o); setShowAutocomplete(false); }} className="block w-full truncate px-3 py-1.5 text-left text-sm hover:bg-blue-50">{o}</button>
                ))}
              </div>
            )}
          </div>
          <Field label="Cutoff Round">
            <select value={refRound} onChange={e => setRefRound(e.target.value)} className="rounded border border-slate-300 bg-white px-2 py-1.5 text-sm">
              <option value="r1">Round I</option><option value="r2">Round II</option><option value="r3">Round III</option><option value="spot">Spot</option>
            </select>
          </Field>
          <span className="text-xs text-slate-500">{totalCount} results</span>
        </div>

        <div className="overflow-hidden rounded border border-slate-300 bg-white">
          <div className="overflow-x-auto">
            <table className="w-full min-w-[820px] text-sm">
              <thead className="sticky top-0 z-10 bg-blue-900 text-left text-xs font-bold uppercase tracking-wide text-white">
                <tr>
                  <th className="px-3 py-2">College / CAP Code</th>
                  <th className="px-3 py-2">Branch</th>
                  <th className="px-3 py-2">District</th>
                  <th className="px-3 py-2">Cat/Seat</th>
                  <th className="px-3 py-2 text-right">Cutoff</th>
                  <th className="px-3 py-2 text-right">Diff</th>
                  <th className="px-3 py-2">Safety</th>
                  <th className="px-3 py-2 text-right">Action</th>
                </tr>
              </thead>
              <tbody>
                {pageRows.map(row => {
                  const d = diffFor(row, refRound);
                  const s = safetyFor(d);
                  const already = preferenceIds.has(row.id);
                  return (
                    <tr key={row.id} className="border-t border-slate-100 hover:bg-slate-50">
                      <td className="px-3 py-2">
                        <button onClick={() => onView(row.id)} className="font-semibold text-blue-800 hover:underline">{row.collegeName}</button>
                        <div className="text-xs text-slate-400">CAP {row.capCode}</div>
                      </td>
                      <td className="px-3 py-2">{row.branchName}</td>
                      <td className="px-3 py-2">{row.district}</td>
                      <td className="px-3 py-2 text-xs">{row.category}/{row.seatType}</td>
                      <td className="px-3 py-2 text-right tabular-nums">{row.cutoffs[refRound] != null ? row.cutoffs[refRound] : '-'}</td>
                      <td className="px-3 py-2 text-right tabular-nums text-slate-500">{d != null ? (d > 0 ? '+' : '') + d : '-'}</td>
                      <td className="px-3 py-2"><SafetyBadge label={s} /></td>
                      <td className="px-3 py-2 text-right">
                        <button disabled={already} onClick={() => onAdd(row.id)} className={`inline-flex items-center gap-1 rounded px-2 py-1 text-xs font-semibold text-white ${already ? 'bg-slate-300' : 'bg-blue-700 hover:bg-blue-800'}`}>
                          <Plus size={12} /> {already ? 'Added' : 'Add'}
                        </button>
                      </td>
                    </tr>
                  );
                })}
                {pageRows.length === 0 && (
                  <tr><td colSpan={8} className="px-3 py-10 text-center text-sm text-slate-400">No colleges match these filters.</td></tr>
                )}
              </tbody>
            </table>
          </div>
          <Pagination page={page} totalPages={totalPages} onChange={setPage} />
        </div>
      </div>
    </div>
  );
}
const BRANCHES_LOOKUP = Object.fromEntries(BRANCHES.map(b => [b.code, b.name]));

/* =========================================================================
   TAB: PREFERENCE LIST
   ========================================================================= */
function PreferencesTab({ preferenceList, rowById, diffFor, safetyFor, refRound, onRemove, onNote, onMove, dragIndexRef, onDrop, onExportJSON, onExportExcel, onPrint, onReset, fileInputRef, onImportJSON, onView }) {
  return (
    <div>
      <div className="mb-3 flex flex-wrap items-center gap-2">
        <button onClick={onExportJSON} className="flex items-center gap-1.5 rounded border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"><Download size={14} /> Export JSON</button>
        <input ref={fileInputRef} type="file" accept=".json" className="hidden" onChange={e => { if (e.target.files[0]) onImportJSON(e.target.files[0]); e.target.value = ''; }} />
        <button onClick={() => fileInputRef.current && fileInputRef.current.click()} className="flex items-center gap-1.5 rounded border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"><Upload size={14} /> Import JSON</button>
        <button onClick={onExportExcel} className="flex items-center gap-1.5 rounded border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"><FileSpreadsheet size={14} /> Export Excel</button>
        <button onClick={onPrint} className="flex items-center gap-1.5 rounded border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"><Printer size={14} /> Print / Save PDF</button>
        <button onClick={onReset} className="ml-auto flex items-center gap-1.5 rounded border border-red-300 bg-white px-3 py-1.5 text-xs font-semibold text-red-600 hover:bg-red-50"><RotateCcw size={14} /> Reset List</button>
      </div>

      {preferenceList.length === 0 ? (
        <div className="rounded border border-dashed border-slate-300 bg-white p-10 text-center text-sm text-slate-400">
          Your preference list is empty. Add colleges from Suggestions or Browse & Filter.
        </div>
      ) : (
        <div className="overflow-hidden rounded border border-slate-300 bg-white">
          <div className="overflow-x-auto">
            <table className="w-full min-w-[860px] text-sm">
              <thead className="sticky top-0 bg-blue-900 text-left text-xs font-bold uppercase tracking-wide text-white">
                <tr>
                  <th className="w-10 px-2 py-2"></th>
                  <th className="px-3 py-2">#</th>
                  <th className="px-3 py-2">College</th>
                  <th className="px-3 py-2">Branch</th>
                  <th className="px-3 py-2 text-right">Cutoff</th>
                  <th className="px-3 py-2 text-right">Diff</th>
                  <th className="px-3 py-2">Safety</th>
                  <th className="px-3 py-2">Notes</th>
                  <th className="px-3 py-2 text-right">Move</th>
                  <th className="px-3 py-2 text-right">Remove</th>
                </tr>
              </thead>
              <tbody>
                {preferenceList.map((p, i) => {
                  const row = rowById.get(p.optionRowId);
                  if (!row) return null;
                  const d = diffFor(row, refRound);
                  const s = safetyFor(d);
                  return (
                    <tr
                      key={p.optionRowId}
                      draggable
                      onDragStart={() => { dragIndexRef.current = i; }}
                      onDragOver={e => e.preventDefault()}
                      onDrop={() => onDrop(i)}
                      className="border-t border-slate-100 hover:bg-slate-50"
                    >
                      <td className="cursor-grab px-2 py-2 text-slate-300"><GripVertical size={16} /></td>
                      <td className="px-3 py-2 font-bold text-blue-900">{i + 1}</td>
                      <td className="px-3 py-2">
                        <button onClick={() => onView(row.id)} className="font-semibold text-blue-800 hover:underline">{row.collegeName}</button>
                        <div className="text-xs text-slate-400">CAP {row.capCode} &middot; {row.category}/{row.seatType}</div>
                      </td>
                      <td className="px-3 py-2">{row.branchName}</td>
                      <td className="px-3 py-2 text-right tabular-nums">{row.cutoffs[refRound] != null ? row.cutoffs[refRound] : '-'}</td>
                      <td className="px-3 py-2 text-right tabular-nums text-slate-500">{d != null ? (d > 0 ? '+' : '') + d : '-'}</td>
                      <td className="px-3 py-2"><SafetyBadge label={s} /></td>
                      <td className="px-3 py-2">
                        <input value={p.notes} onChange={e => onNote(p.optionRowId, e.target.value)} placeholder="Add a note…" className="w-36 rounded border border-slate-200 px-2 py-1 text-xs focus:border-blue-500 focus:outline-none" />
                      </td>
                      <td className="px-3 py-2 text-right">
                        <button onClick={() => onMove(i, -1)} disabled={i === 0} className="rounded border border-slate-200 p-1 disabled:opacity-30"><ChevronUp size={14} /></button>
                        <button onClick={() => onMove(i, 1)} disabled={i === preferenceList.length - 1} className="ml-1 rounded border border-slate-200 p-1 disabled:opacity-30"><ChevronDown size={14} /></button>
                      </td>
                      <td className="px-3 py-2 text-right">
                        <button onClick={() => onRemove(p.optionRowId)} className="rounded border border-red-200 p-1 text-red-500 hover:bg-red-50"><Trash2 size={14} /></button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

/* =========================================================================
   TAB: COMPARE
   ========================================================================= */
function RowPicker({ label, value, onChange, dataset }) {
  const [q, setQ] = useState('');
  const [open, setOpen] = useState(false);
  const options = useMemo(() => {
    const s = q.trim().toLowerCase();
    if (!s) return [];
    return dataset.filter(r => `${r.collegeName} ${r.branchName}`.toLowerCase().includes(s)).slice(0, 8);
  }, [q, dataset]);
  return (
    <div className="relative">
      <Field label={label}>
        <input
          value={value ? `${value.collegeName} — ${value.branchName}` : q}
          onChange={e => { setQ(e.target.value); onChange(null); setOpen(true); }}
          onFocus={() => setOpen(true)}
          onBlur={() => setTimeout(() => setOpen(false), 150)}
          placeholder="Search college + branch…"
          className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-sm focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600"
        />
      </Field>
      {open && options.length > 0 && (
        <div className="absolute z-10 mt-1 w-full rounded border border-slate-300 bg-white shadow-lg">
          {options.map(o => (
            <button key={o.id} onMouseDown={() => { onChange(o); setQ(''); setOpen(false); }} className="block w-full truncate px-3 py-1.5 text-left text-xs hover:bg-blue-50">
              {o.collegeName} — {o.branchName} <span className="text-slate-400">({o.category}/{o.seatType})</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function CompareTab({ dataset, compareA, compareB, setCompareA, setCompareB, collegeMeta, refRound, diffFor, safetyFor }) {
  const rowA = compareA != null ? dataset.find(r => r.id === compareA) : null;
  const rowB = compareB != null ? dataset.find(r => r.id === compareB) : null;

  const rows = [
    ['CAP Code', r => r.capCode],
    ['District', r => r.district],
    ['University', r => r.university],
    ['Autonomy', r => r.autonomy],
    ['Type', r => r.instType],
    ['NAAC Grade', r => r.naac || '—'],
    ['Branch', r => r.branchName],
    ['Category / Seat', r => `${r.category} / ${r.seatType}`],
    ['Round I Cutoff', r => r.cutoffs.r1 != null ? r.cutoffs.r1 : '-'],
    ['Round II Cutoff', r => r.cutoffs.r2 != null ? r.cutoffs.r2 : '-'],
    ['Round III Cutoff', r => r.cutoffs.r3 != null ? r.cutoffs.r3 : '-'],
    ['Spot Round Cutoff', r => r.cutoffs.spot != null ? r.cutoffs.spot : '-'],
    ['Difference (selected round)', r => { const d = diffFor(r, refRound); return d != null ? d : '-'; }],
    ['Safety Label', r => safetyFor(diffFor(r, refRound)) || '—'],
    ['Avg. Placement (sample)', r => `₹${collegeMeta[r.collegeName] ? collegeMeta[r.collegeName].avgPackageLPA : '-'} LPA`],
    ['Fees / Year (sample)', r => `₹${collegeMeta[r.collegeName] ? collegeMeta[r.collegeName].feesPerYear.toLocaleString('en-IN') : '-'}`],
  ];

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
        <RowPicker label="College A" value={rowA} onChange={setCompareA ? (r) => setCompareA(r ? r.id : null) : () => {}} dataset={dataset} />
        <RowPicker label="College B" value={rowB} onChange={setCompareB ? (r) => setCompareB(r ? r.id : null) : () => {}} dataset={dataset} />
      </div>

      {(!rowA || !rowB) ? (
        <div className="rounded border border-dashed border-slate-300 bg-white p-10 text-center text-sm text-slate-400">Choose two college + branch combinations to compare them side by side.</div>
      ) : (
        <div className="overflow-x-auto rounded border border-slate-300 bg-white">
          <table className="w-full min-w-[560px] text-sm">
            <thead className="bg-blue-900 text-left text-xs font-bold uppercase tracking-wide text-white">
              <tr><th className="px-3 py-2">Attribute</th><th className="px-3 py-2">{rowA.collegeName}</th><th className="px-3 py-2">{rowB.collegeName}</th></tr>
            </thead>
            <tbody>
              {rows.map(([label, fn]) => (
                <tr key={label} className="border-t border-slate-100">
                  <td className="px-3 py-2 font-semibold text-slate-600">{label}</td>
                  <td className="px-3 py-2">{String(fn(rowA))}</td>
                  <td className="px-3 py-2">{String(fn(rowB))}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* =========================================================================
   TAB: STATISTICS
   ========================================================================= */
function StatCard({ label, value }) {
  return (
    <div className="rounded border border-slate-300 bg-white p-4">
      <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</div>
      <div className="mt-1 text-2xl font-bold text-blue-900">{value == null ? '—' : value}</div>
    </div>
  );
}

function StatsTab({ stats, refRound }) {
  const order = ['Dream', 'Reach', 'Competitive', 'Safe', 'Very Safe'];
  const colorBar = { 'Dream': 'bg-red-400', 'Reach': 'bg-orange-400', 'Competitive': 'bg-amber-400', 'Safe': 'bg-green-500', 'Very Safe': 'bg-teal-500' };
  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
        <StatCard label="Colleges Selected" value={stats.totalColleges} />
        <StatCard label="Branches Selected" value={stats.totalBranches} />
        <StatCard label="Total Entries" value={stats.total} />
        <StatCard label="Highest Cutoff" value={stats.highest} />
        <StatCard label="Lowest Cutoff" value={stats.lowest} />
        <StatCard label="Average Cutoff" value={stats.average} />
      </div>
      <div className="rounded border border-slate-300 bg-white p-4">
        <h3 className="mb-3 text-sm font-bold uppercase tracking-wide text-blue-900">Safety Distribution ({refRound === 'r1' ? 'Round I' : refRound === 'r2' ? 'Round II' : refRound === 'r3' ? 'Round III' : 'Spot'})</h3>
        {stats.total === 0 ? (
          <p className="text-sm text-slate-400">Add colleges to your preference list to see this breakdown.</p>
        ) : (
          <div className="space-y-2">
            {order.map(label => {
              const c = stats.counts[label];
              const pct = stats.total ? Math.round((c / stats.total) * 100) : 0;
              return (
                <div key={label} className="flex items-center gap-3">
                  <span className="w-28 shrink-0 text-xs font-semibold text-slate-600">{label}</span>
                  <div className="h-3 flex-1 rounded bg-slate-100">
                    <div className={`h-3 rounded ${colorBar[label]}`} style={{ width: `${pct}%` }} />
                  </div>
                  <span className="w-14 shrink-0 text-right text-xs tabular-nums text-slate-500">{c} ({pct}%)</span>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

/* =========================================================================
   COLLEGE DETAIL MODAL
   ========================================================================= */
function DetailModal({ row, meta, onClose, onAdd, onCompare }) {
  const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(row.collegeName + ' official website')}`;
  return (
    <Modal title={row.collegeName} onClose={onClose} wide>
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
        <div>
          <h4 className="mb-1 text-xs font-bold uppercase tracking-wide text-slate-500">Location & Affiliation</h4>
          <p className="text-sm">{row.district} &middot; {row.university}</p>
          <p className="text-sm text-slate-600">{row.autonomy} &middot; {row.instType}{row.naac ? ` · NAAC ${row.naac}` : ''}</p>
          {row.minority !== 'None' && <p className="text-sm text-slate-600">Minority status: {row.minority}</p>}
        </div>
        <div>
          <h4 className="mb-1 text-xs font-bold uppercase tracking-wide text-slate-500">This Option Row</h4>
          <p className="text-sm">{row.branchName}</p>
          <p className="text-sm text-slate-600">Category {row.category} &middot; Seat Type {row.seatType} &middot; CAP Code {row.capCode}</p>
        </div>
      </div>

      <h4 className="mb-1 mt-4 text-xs font-bold uppercase tracking-wide text-slate-500">Previous Year Closing Percentiles (sample)</h4>
      <table className="w-full text-sm">
        <thead><tr className="border-b border-slate-200 text-left text-xs text-slate-500"><th className="py-1">Round I</th><th className="py-1">Round II</th><th className="py-1">Round III</th><th className="py-1">Spot</th></tr></thead>
        <tbody><tr>
          <td className="py-1">{row.cutoffs.r1 != null ? row.cutoffs.r1 : '-'}</td>
          <td className="py-1">{row.cutoffs.r2 != null ? row.cutoffs.r2 : '-'}</td>
          <td className="py-1">{row.cutoffs.r3 != null ? row.cutoffs.r3 : '-'}</td>
          <td className="py-1">{row.cutoffs.spot != null ? row.cutoffs.spot : 'Not held'}</td>
        </tr></tbody>
      </table>

      {meta && (
        <>
          <h4 className="mb-1 mt-4 text-xs font-bold uppercase tracking-wide text-slate-500">Seat Matrix & Placements (illustrative)</h4>
          <p className="text-sm text-slate-600">Approx. intake: {meta.intake} seats/branch &middot; Avg. package: ₹{meta.avgPackageLPA} LPA &middot; Fees: ₹{meta.feesPerYear.toLocaleString('en-IN')}/year</p>
        </>
      )}

      <div className="mt-5 flex flex-wrap items-center gap-2">
        <button onClick={onAdd} className="flex items-center gap-1.5 rounded bg-blue-700 px-3 py-1.5 text-sm font-semibold text-white hover:bg-blue-800"><Plus size={14} /> Add to Preferences</button>
        <button onClick={onCompare} className="flex items-center gap-1.5 rounded border border-slate-300 px-3 py-1.5 text-sm font-semibold text-slate-700 hover:bg-slate-50"><ArrowLeftRight size={14} /> Compare</button>
        <a href={searchUrl} target="_blank" rel="noreferrer" className="ml-auto flex items-center gap-1.5 text-sm font-semibold text-blue-700 hover:underline"><ExternalLink size={14} /> Find official website</a>
      </div>
    </Modal>
  );
}
