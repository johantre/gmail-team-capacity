// ============================================================
// CONFIG — constants shared across the project
// ============================================================

const WERKUREN_PER_DAG = 8;

// Inline SVG icons (white, Tabler style) — no internet needed, sourced from res/
const SVG_CALENDAR = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5m0 2a2 2 0 0 1 2 -2h12a2 2 0 0 1 2 2v12a2 2 0 0 1 -2 2h-12a2 2 0 0 1 -2 -2z"/><path d="M16 3l0 4"/><path d="M8 3l0 4"/><path d="M4 11l16 0"/><path d="M8 15h2v2h-2z"/></svg>';
const SVG_CHART    = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M7 16v-5"/><path d="M11 16v-8"/><path d="M15 16v-3"/></svg>';
const SHEET_TEAMS = "Teams";
const SHEET_CAPACITEIT = "Capacity";

// Acme brand colors
const BLAUW      = "#3E7CBF";
const TURQUOISE  = "#27B4AF";
const GRIJS      = "#3C3C3B";
const BLAUW_LICHT     = "#d6e8f5";
const TURQUOISE_LICHT = "#d0edec";

const TEAM_KLEUREN = {
  "Team A": BLAUW,
  "Team B": TURQUOISE,
  "Team C": "#2e6da4", // darker blue
  "Team D": "#1e9490", // darker turquoise
};

const MAAND_MAP = {
  jan:0, feb:1, mar:2, apr:3, may:4, jun:5,
  jul:6, aug:7, sep:8, oct:9, nov:10, dec:11,
  mrt:2, mei:4, okt:9,
};
