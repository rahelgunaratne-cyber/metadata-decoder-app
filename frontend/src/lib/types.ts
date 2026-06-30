export type Row = Record<string, unknown>;

export interface Counts {
  artistTypos: number;
  isrcConflicts: number;
  isrcConflictGroups: number;
  missingFields: number;
  formatIssues: number;
  splitErrors: number;
  idMismatches: number;
  total: number;
}

export interface Scan {
  id: string;
  filename: string;
  created_at: string;
  updated_at: string;
  uploaded_by: string;
  tracks_sheet: string;
  sheets_scanned: string[];
  detected_format: string;
  other_sheets: string[];
  is_rescan: boolean;
  status: string;
  counts: Counts;
  keys: Record<string, string>;
}

export interface Results {
  stats: Row;
  issues: Row[];
  artistClusters: Row[];
  isrcConflicts: Row[];
  missingSummary: Row[];
  missingCells: Row[];
  formatColumns: Row[];
  formatRows: Row[];
  formatCorrections: Row[];
  splitsReview: Row[];
  splitErrors: Row[];
  idMismatches: Row[];
  detectedFormat: string;
}

export interface ScanDetail {
  scan: Scan;
  results: Results;
}

export interface AppConfig {
  authEnabled: boolean;
  oauthClientId: string;
  allowedDomain: string;
}

export interface AppUser {
  email: string;
  name: string;
  picture: string;
}
