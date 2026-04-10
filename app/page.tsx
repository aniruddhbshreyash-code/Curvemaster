'use client';

import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { Upload, Settings, BarChart3, Download, X, CheckCircle2, AlertCircle, Users, AlertTriangle } from 'lucide-react';
import * as XLSX from 'xlsx';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Cell, Legend } from 'recharts';

// Types
interface StudentData {
  [key: string]: string | number;
}

interface ColumnConfig {
  columnName: string;
  maxMarks: number | string;
  weightage: number | string;
}

interface TeamData {
  teamName: string;
  members: number[];
  marks: number;
}

interface ProcessedStudent {
  name: string;
  finalScore: number;
  grade: string;
  rawData: StudentData;
  individualComponent?: number;
  teamComponent?: number;
  teamName?: string;
  teamRawScore?: number;
}

interface GradeCutoffs {
  A: number;
  'A-': number;
  B: number;
  'B-': number;
  C: number;
  'C-': number;
  D: number;
}

interface TeamMatchSummary {
  studentsInTeams: number;
  studentsWithoutTeams: number;
  totalTeams: number;
  unmatchedMembers: number[];
  duplicateMembers: number[];
}

type AppState = 'upload' | 'configure' | 'dashboard';

interface HistogramTooltipPayload {
  dataKey: string;
  value: number;
  color: string;
  name: string;
}

interface HistogramTooltipProps {
  active?: boolean;
  payload?: HistogramTooltipPayload[];
  label?: number;
}

const GRADE_COLORS = {
  'A': '#059669',
  'A-': '#10b981',
  'B': '#2563eb',
  'B-': '#3b82f6',
  'C': '#d97706',
  'C-': '#f59e0b',
  'D': '#ea580c',
  'F': '#dc2626',
};

const GRADE_LABELS: Record<string, string> = {
  'A': 'Excellent',
  'A-': 'Very Good',
  'B': 'Good',
  'B-': 'Above Average',
  'C': 'Average',
  'C-': 'Below Average',
  'D': 'Poor',
  'F': 'Fail',
};

// Custom tooltip for the histogram
const HistogramChartTooltip = ({ active, payload, label }: HistogramTooltipProps) => {
  if (!active || !payload || !payload.length) return null;

  const gradeEntries = payload.filter(p => p.value > 0);
  if (gradeEntries.length === 0) return null;

  const totalCount = gradeEntries.reduce((sum, p) => sum + p.value, 0);

  return (
    <div style={{
      backgroundColor: 'white',
      border: '1px solid #e2e8f0',
      borderRadius: '8px',
      boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)',
      padding: '12px 16px',
      fontSize: '13px',
    }}>
      <p style={{ color: '#0f172a', fontWeight: 'bold', marginBottom: '6px' }}>Score: {label}</p>
      {gradeEntries.map((entry, i) => (
        <p key={i} style={{ color: '#0f172a', margin: '2px 0', fontWeight: 500 }}>
          {entry.name}: {entry.value} student{entry.value !== 1 ? 's' : ''}
        </p>
      ))}
      <hr style={{ margin: '4px 0', borderColor: '#e2e8f0' }} />
      <p style={{ color: '#0f172a', fontWeight: 600, margin: '2px 0' }}>
        Total: {totalCount} student{totalCount !== 1 ? 's' : ''}
      </p>
    </div>
  );
};

export default function CurveMaster() {
  const [mounted, setMounted] = useState(false);
  const [state, setState] = useState<AppState>('upload');
  const [rawData, setRawData] = useState<StudentData[]>([]);
  const [fileName, setFileName] = useState<string>('');
  
  // Configuration
  const [numericColumns, setNumericColumns] = useState<string[]>([]);
  const [selectedColumns, setSelectedColumns] = useState<Set<string>>(new Set());
  const [columnConfigs, setColumnConfigs] = useState<Map<string, ColumnConfig>>(new Map());
  const [nameColumn, setNameColumn] = useState<string>('');
  const [allColumns, setAllColumns] = useState<string[]>([]);
  
  // Processed data
  const [processedData, setProcessedData] = useState<ProcessedStudent[]>([]);
  const [mean, setMean] = useState<number>(0);
  const [stdDev, setStdDev] = useState<number>(0);
  const [cutoffs, setCutoffs] = useState<GradeCutoffs>({
    A: 0, 'A-': 0, B: 0, 'B-': 0, C: 0, 'C-': 0, D: 0
  });

  // ─── Team Project State ───
  const [includeTeamProjects, setIncludeTeamProjects] = useState(false);
  const [teamWeightage, setTeamWeightage] = useState<number>(30);
  const [teamRawData, setTeamRawData] = useState<StudentData[]>([]);
  const [teamFileName, setTeamFileName] = useState<string>('');
  const [teamColumns, setTeamColumns] = useState<string[]>([]);
  const [teamMarksColumn, setTeamMarksColumn] = useState<string>('');
  const [maxTeamMarks, setMaxTeamMarks] = useState<number>(100);
  const [rollNumberColumn, setRollNumberColumn] = useState<string>('');
  const [teamMatchSummary, setTeamMatchSummary] = useState<TeamMatchSummary | null>(null);
  const [teamValidationErrors, setTeamValidationErrors] = useState<string[]>([]);

  useEffect(() => {
    setMounted(true);
  }, []);

  // File upload handler (individual students)
  const handleFileUpload = useCallback((file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet) as StudentData[];
        
        if (json.length === 0) {
          alert('The file appears to be empty');
          return;
        }

        setRawData(json);
        setFileName(file.name);
        
        // Detect columns
        const columns = Object.keys(json[0]);
        setAllColumns(columns);
        
        // Detect numeric columns
        const numeric = columns.filter(col => {
          return json.every(row => {
            const val = row[col];
            return val === null || val === undefined || val === '' || !isNaN(Number(val));
          });
        });
        
        setNumericColumns(numeric);

        // Go to configure immediately after file read
        setState('configure');
      } catch (error) {
        alert('Error reading file. Please ensure it\'s a valid CSV or Excel file.');
        console.error(error);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  // Team file upload handler
  const handleTeamFileUpload = useCallback((file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet) as StudentData[];

        if (json.length === 0) {
          alert('The team file appears to be empty');
          return;
        }

        setTeamRawData(json);
        setTeamFileName(file.name);
        const columns = Object.keys(json[0]);
        setTeamColumns(columns);

        // Auto-detect marks column (often named "Marks" or similar)
        const marksCol = columns.find(c => c.toLowerCase().includes('mark'));
        if (marksCol) setTeamMarksColumn(marksCol);
      } catch (error) {
        alert('Error reading team file. Please ensure it\'s a valid CSV or Excel file.');
        console.error(error);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  // Proceed to configure from upload
  const proceedToConfigure = useCallback(() => {
    if (rawData.length === 0) {
      alert('Please upload the individual students file first.');
      return;
    }
    setState('configure');
  }, [rawData]);

  // Parse team data and build member→team map
  const parseTeamData = useCallback((): { teams: TeamData[]; memberToTeam: Map<number, { teamName: string; marks: number }> } => {
    if (teamRawData.length === 0 || !teamMarksColumn) return { teams: [], memberToTeam: new Map() };

    const teamNameCol = teamColumns.find(c => c.toLowerCase().includes('team'));
    const memberCols = teamColumns.filter(c => c.toLowerCase().includes('member'));

    const teams: TeamData[] = [];
    const memberToTeam = new Map<number, { teamName: string; marks: number }>();

    teamRawData.forEach(row => {
      const teamName = String(row[teamNameCol || teamColumns[0]] || 'Unknown');
      const marks = Number(row[teamMarksColumn]) || 0;
      const members: number[] = [];

      memberCols.forEach(col => {
        const val = row[col];
        if (val !== null && val !== undefined && val !== '') {
          const num = Number(val);
          if (!isNaN(num)) {
            members.push(num);
            memberToTeam.set(num, { teamName, marks });
          }
        }
      });

      teams.push({ teamName, members, marks });
    });

    return { teams, memberToTeam };
  }, [teamRawData, teamMarksColumn, teamColumns]);

  // Extract member number from roll number
  const extractMemberNumber = useCallback((rollNumber: string): number | null => {
    const str = String(rollNumber).trim();
    // Extract last 3+ digits from the roll number
    const match = str.match(/(\d{1,3})$/);
    if (match) {
      return parseInt(match[1], 10);
    }
    return null;
  }, []);

  // Validate team matching
  const validateTeams = useCallback(() => {
    if (!includeTeamProjects || teamRawData.length === 0 || !rollNumberColumn || !teamMarksColumn) {
      setTeamMatchSummary(null);
      setTeamValidationErrors([]);
      return;
    }

    const { teams, memberToTeam } = parseTeamData();
    const errors: string[] = [];

    // Check for duplicate members across teams
    const allMembers: number[] = [];
    const duplicates: number[] = [];
    teams.forEach(team => {
      team.members.forEach(m => {
        if (allMembers.includes(m)) {
          duplicates.push(m);
        }
        allMembers.push(m);
      });
    });

    if (duplicates.length > 0) {
      errors.push(`Duplicate member numbers found in multiple teams: ${[...new Set(duplicates)].join(', ')}`);
    }

    // Match students to teams
    let studentsInTeams = 0;
    let studentsWithoutTeams = 0;
    const unmatchedMembers: number[] = [];

    rawData.forEach(student => {
      const rollNumber = String(student[rollNumberColumn] || '');
      const memberNum = extractMemberNumber(rollNumber);
      if (memberNum !== null && memberToTeam.has(memberNum)) {
        studentsInTeams++;
      } else {
        studentsWithoutTeams++;
      }
    });

    // Check which team members don't match any student
    const studentMemberNumbers = new Set(
      rawData.map(s => extractMemberNumber(String(s[rollNumberColumn] || ''))).filter((n): n is number => n !== null)
    );
    allMembers.forEach(m => {
      if (!studentMemberNumbers.has(m)) {
        unmatchedMembers.push(m);
      }
    });

    if (unmatchedMembers.length > 0) {
      errors.push(`Team member numbers not found in student list: ${unmatchedMembers.join(', ')}`);
    }

    setTeamMatchSummary({
      studentsInTeams,
      studentsWithoutTeams,
      totalTeams: teams.length,
      unmatchedMembers,
      duplicateMembers: [...new Set(duplicates)],
    });
    setTeamValidationErrors(errors);
  }, [includeTeamProjects, teamRawData, rollNumberColumn, teamMarksColumn, rawData, parseTeamData, extractMemberNumber]);

  // Run validation whenever relevant config changes
  useEffect(() => {
    if (state === 'configure') {
      validateTeams();
    }
  }, [state, rollNumberColumn, teamMarksColumn, validateTeams]);

  // Column selection handler
  const handleColumnToggle = (column: string) => {
    const newSelected = new Set(selectedColumns);
    if (newSelected.has(column)) {
      newSelected.delete(column);
      const newConfigs = new Map(columnConfigs);
      newConfigs.delete(column);
      setColumnConfigs(newConfigs);
    } else {
      newSelected.add(column);
      // Initialize with default values
      const newConfigs = new Map(columnConfigs);
      newConfigs.set(column, {
        columnName: column,
        maxMarks: 100,
        weightage: 100 / (selectedColumns.size + 1)
      });
      setColumnConfigs(newConfigs);
    }
    setSelectedColumns(newSelected);
  };

  const updateColumnConfig = (column: string, field: 'maxMarks' | 'weightage', value: string) => {
    const newConfigs = new Map(columnConfigs);
    const config = newConfigs.get(column);
    if (config) {
      // Allow empty string temporarily during editing
      config[field] = value === '' ? '' : value;
      newConfigs.set(column, config);
      setColumnConfigs(newConfigs);
    }
  };

  const handleConfigBlur = (column: string, field: 'maxMarks' | 'weightage') => {
    const newConfigs = new Map(columnConfigs);
    const config = newConfigs.get(column);
    if (config) {
      const currentValue = config[field];
      // Convert empty string or invalid values to defaults on blur
      if (currentValue === '' || currentValue === null || currentValue === undefined) {
        config[field] = field === 'maxMarks' ? 100 : 0;
      } else {
        // Ensure it's a valid number
        const numValue = Number(currentValue);
        config[field] = isNaN(numValue) ? (field === 'maxMarks' ? 100 : 0) : numValue;
      }
      newConfigs.set(column, config);
      setColumnConfigs(newConfigs);
    }
  };

  // Calculate population standard deviation (numpy.std with ddof=0)
  const calculatePopulationStdDev = (values: number[]): number => {
    const n = values.length;
    const avg = values.reduce((a, b) => a + b, 0) / n;
    const squaredDiffs = values.map(val => Math.pow(val - avg, 2));
    const variance = squaredDiffs.reduce((a, b) => a + b, 0) / n; // ddof=0
    return Math.sqrt(variance);
  };

  // Check if team feature has blocking errors (duplicates)
  const hasTeamBlockingErrors = useMemo(() => {
    if (!includeTeamProjects) return false;
    return (teamMatchSummary?.duplicateMembers.length ?? 0) > 0;
  }, [includeTeamProjects, teamMatchSummary]);

  // Process data and calculate grades
  const processData = useCallback(() => {
    if (selectedColumns.size === 0) {
      alert('Please select at least one column to grade');
      return;
    }

    if (!nameColumn) {
      alert('Please select a name column');
      return;
    }

    if (includeTeamProjects && hasTeamBlockingErrors) {
      alert('Please fix team data errors before proceeding (duplicate members found).');
      return;
    }

    if (includeTeamProjects && !rollNumberColumn) {
      alert('Please select a Roll Number column for team matching.');
      return;
    }

    if (includeTeamProjects && !teamMarksColumn) {
      alert('Please select a Team Marks column.');
      return;
    }

    // Build team map if team projects enabled
    let memberToTeam: Map<number, { teamName: string; marks: number }> = new Map();
    if (includeTeamProjects && teamRawData.length > 0) {
      const parsed = parseTeamData();
      memberToTeam = parsed.memberToTeam;
    }

    const teamEnabled = includeTeamProjects && memberToTeam.size > 0;
    const individualWeightFactor = teamEnabled ? (100 - teamWeightage) / 100 : 1;
    const teamWeightFactor = teamEnabled ? teamWeightage / 100 : 0;

    // Calculate final scores
    const studentsWithScores = rawData.map(student => {
      // Individual component
      let individualSum = 0;
      selectedColumns.forEach(col => {
        const config = columnConfigs.get(col);
        if (config) {
          const rawScore = Number(student[col]) || 0;
          const maxMarks = Number(config.maxMarks) || 100;
          const weightage = Number(config.weightage) || 0;
          const normalized = (rawScore / maxMarks) * weightage;
          individualSum += normalized;
        }
      });

      const individualComponent = individualSum * individualWeightFactor;

      // Team component
      let teamComponent = 0;
      let teamName: string | undefined = undefined;
      let teamRawScore: number | undefined = undefined;

      if (teamEnabled && rollNumberColumn) {
        const rollNumber = String(student[rollNumberColumn] || '');
        const memberNum = extractMemberNumber(rollNumber);
        if (memberNum !== null) {
          const teamInfo = memberToTeam.get(memberNum);
          if (teamInfo) {
            teamName = teamInfo.teamName;
            teamRawScore = teamInfo.marks;
            const normalizedTeam = (teamInfo.marks / (maxTeamMarks || 100)) * 100;
            teamComponent = normalizedTeam * teamWeightFactor;
          }
        }
      }

      const finalScore = individualComponent + teamComponent;

      return {
        name: String(student[nameColumn]),
        finalScore,
        grade: '',
        rawData: student,
        ...(teamEnabled ? { individualComponent, teamComponent, teamName, teamRawScore } : {}),
      } as ProcessedStudent;
    });

    // Calculate statistics
    const scores = studentsWithScores.map(s => s.finalScore);
    const meanScore = scores.reduce((a, b) => a + b, 0) / scores.length;
    const stdDevScore = calculatePopulationStdDev(scores);

    setMean(meanScore);
    setStdDev(stdDevScore);

    // Calculate initial cutoffs
    const initialCutoffs: GradeCutoffs = {
      'A': meanScore + 1.5 * stdDevScore,
      'A-': meanScore + 1.0 * stdDevScore,
      'B': meanScore + 0.5 * stdDevScore,
      'B-': meanScore,
      'C': meanScore - 0.5 * stdDevScore,
      'C-': meanScore - 1.0 * stdDevScore,
      'D': meanScore - 1.5 * stdDevScore,
    };

    setCutoffs(initialCutoffs);

    // Assign grades
    const assignGrade = (score: number, cutoffs: GradeCutoffs): string => {
      if (score >= cutoffs.A) return 'A';
      if (score >= cutoffs['A-']) return 'A-';
      if (score >= cutoffs.B) return 'B';
      if (score >= cutoffs['B-']) return 'B-';
      if (score >= cutoffs.C) return 'C';
      if (score >= cutoffs['C-']) return 'C-';
      if (score >= cutoffs.D) return 'D';
      return 'F';
    };

    const graded = studentsWithScores.map(s => ({
      ...s,
      grade: assignGrade(s.finalScore, initialCutoffs)
    }));

    // Sort by score descending
    graded.sort((a, b) => b.finalScore - a.finalScore);

    setProcessedData(graded);
    setState('dashboard');
  }, [rawData, selectedColumns, columnConfigs, nameColumn, includeTeamProjects, hasTeamBlockingErrors, rollNumberColumn, teamMarksColumn, teamRawData, parseTeamData, teamWeightage, maxTeamMarks, extractMemberNumber]);

  // Update grades when cutoffs change
  useEffect(() => {
    const assignGrade = (score: number): string => {
      if (score >= cutoffs.A) return 'A';
      if (score >= cutoffs['A-']) return 'A-';
      if (score >= cutoffs.B) return 'B';
      if (score >= cutoffs['B-']) return 'B-';
      if (score >= cutoffs.C) return 'C';
      if (score >= cutoffs['C-']) return 'C-';
      if (score >= cutoffs.D) return 'D';
      return 'F';
    };

    setProcessedData(current =>
      current.map(student => ({
        ...student,
        grade: assignGrade(student.finalScore)
      }))
    );
  }, [cutoffs]);

  // Whether team data is active in the current processed results
  const teamEnabledInResults = useMemo(() => {
    return processedData.length > 0 && processedData[0].individualComponent !== undefined;
  }, [processedData]);

  // Team stats for dashboard
  const teamDashboardStats = useMemo(() => {
    if (!teamEnabledInResults) return null;
    const inTeams = processedData.filter(s => s.teamName).length;
    return { inTeams, total: processedData.length };
  }, [teamEnabledInResults, processedData]);

  // Export modal state
  const [showExportModal, setShowExportModal] = useState(false);
  const [exportSelectedColumns, setExportSelectedColumns] = useState<Set<string>>(new Set());

  // Build available export columns
  const availableExportColumns = useMemo(() => {
    const cols: string[] = [];
    if (processedData.length > 0) {
      // All rawData keys
      const rawKeys = Object.keys(processedData[0].rawData);
      cols.push(...rawKeys);
    }
    // Conditional team columns
    if (teamEnabledInResults) {
      cols.push('Team Name', 'Team Project Score', 'Individual Component', 'Team Component');
    }
    // Always
    cols.push('Final Score', 'Grade');
    return cols;
  }, [processedData, teamEnabledInResults]);

  // Initialize export columns when modal opens
  const openExportModal = useCallback(() => {
    setExportSelectedColumns(new Set(availableExportColumns));
    setShowExportModal(true);
  }, [availableExportColumns]);

  const toggleExportColumn = (col: string) => {
    setExportSelectedColumns(prev => {
      const next = new Set(prev);
      if (next.has(col)) {
        next.delete(col);
      } else {
        next.add(col);
      }
      return next;
    });
  };

  // Export to Excel with selected columns
  const exportToExcel = () => {
    const exportData = processedData.map(student => {
      const full: Record<string, string | number> = { ...student.rawData };

      if (teamEnabledInResults) {
        full['Team Name'] = student.teamName || 'No Team';
        full['Team Project Score'] = student.teamRawScore ?? 'N/A';
        full['Individual Component'] = student.individualComponent?.toFixed(2) ?? '';
        full['Team Component'] = student.teamComponent?.toFixed(2) ?? '0.00';
      }

      full['Final Score'] = student.finalScore.toFixed(2);
      full['Grade'] = student.grade;

      // Filter to only selected columns
      const filtered: Record<string, string | number> = {};
      availableExportColumns.forEach(col => {
        if (exportSelectedColumns.has(col) && full[col] !== undefined) {
          filtered[col] = full[col];
        }
      });
      return filtered;
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Graded Results');
    XLSX.writeFile(wb, `graded_${fileName}`);
    setShowExportModal(false);
  };

  const reset = () => {
    setState('upload');
    setRawData([]);
    setFileName('');
    setSelectedColumns(new Set());
    setColumnConfigs(new Map());
    setProcessedData([]);
    // Reset team state
    setIncludeTeamProjects(false);
    setTeamWeightage(30);
    setTeamRawData([]);
    setTeamFileName('');
    setTeamColumns([]);
    setTeamMarksColumn('');
    setMaxTeamMarks(100);
    setRollNumberColumn('');
    setTeamMatchSummary(null);
    setTeamValidationErrors([]);
  };

  if (!mounted) {
    return null;
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-blue-50 to-indigo-50">
      {/* Header */}
      <header className="border-b border-slate-200 bg-white/80 backdrop-blur-sm sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-gradient-to-br from-blue-600 to-indigo-600 rounded-lg flex items-center justify-center">
                <BarChart3 className="w-6 h-6 text-white" />
              </div>
              <div>
                <h1 className="text-2xl font-bold bg-gradient-to-r from-blue-600 to-indigo-600 bg-clip-text text-transparent">
                  CurveMaster
                </h1>
                <p className="text-xs text-slate-500">Relative Grading System</p>
              </div>
            </div>
            {state !== 'upload' && (
              <button
                onClick={reset}
                className="flex items-center gap-2 px-4 py-2 text-sm text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded-lg transition-colors"
              >
                <X className="w-4 h-4" />
                Start Over
              </button>
            )}
          </div>
        </div>
      </header>

      {/* Progress Indicator */}
      <div className="max-w-7xl mx-auto px-6 py-8">
        <div className="flex items-center justify-center gap-4">
          <div className={`flex items-center gap-2 ${state === 'upload' ? 'text-blue-600' : 'text-green-600'}`}>
            {state === 'upload' ? <div className="w-8 h-8 rounded-full border-2 border-blue-600 flex items-center justify-center text-sm font-semibold">1</div> : <CheckCircle2 className="w-8 h-8" />}
            <span className="text-sm font-medium">Upload</span>
          </div>
          <div className={`h-0.5 w-24 ${state === 'upload' ? 'bg-slate-200' : 'bg-green-600'}`} />
          <div className={`flex items-center gap-2 ${state === 'configure' ? 'text-blue-600' : state === 'dashboard' ? 'text-green-600' : 'text-slate-400'}`}>
            {state === 'dashboard' ? <CheckCircle2 className="w-8 h-8" /> : <div className={`w-8 h-8 rounded-full border-2 ${state === 'configure' ? 'border-blue-600' : 'border-slate-300'} flex items-center justify-center text-sm font-semibold`}>2</div>}
            <span className="text-sm font-medium">Configure</span>
          </div>
          <div className={`h-0.5 w-24 ${state === 'dashboard' ? 'bg-green-600' : 'bg-slate-200'}`} />
          <div className={`flex items-center gap-2 ${state === 'dashboard' ? 'text-blue-600' : 'text-slate-400'}`}>
            <div className={`w-8 h-8 rounded-full border-2 ${state === 'dashboard' ? 'border-blue-600' : 'border-slate-300'} flex items-center justify-center text-sm font-semibold`}>3</div>
            <span className="text-sm font-medium">Dashboard</span>
          </div>
        </div>
      </div>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-6 pb-12">
        {/* ═══════════════════ Upload State ═══════════════════ */}
        {state === 'upload' && (
          <div className="max-w-2xl mx-auto space-y-6">
            {/* Main file upload */}
            <div className="bg-white rounded-2xl shadow-lg shadow-blue-100/50 border border-slate-200 p-12">
              <div className="text-center mb-8">
                <h2 className="text-2xl font-bold text-slate-900 mb-2">Upload Student Data</h2>
                <p className="text-slate-600">Drop your CSV or Excel file to get started</p>
              </div>
              
              <div
                className={`border-2 border-dashed rounded-xl p-12 text-center hover:border-blue-500 hover:bg-blue-50/50 transition-all cursor-pointer group ${fileName ? 'border-green-400 bg-green-50/30' : 'border-blue-300'}`}
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                  e.preventDefault();
                  const file = e.dataTransfer.files[0];
                  if (file) handleFileUpload(file);
                }}
                onClick={() => {
                  const input = document.createElement('input');
                  input.type = 'file';
                  input.accept = '.csv,.xlsx,.xls';
                  input.onchange = (e) => {
                    const file = (e.target as HTMLInputElement).files?.[0];
                    if (file) handleFileUpload(file);
                  };
                  input.click();
                }}
              >
                {fileName ? (
                  <>
                    <CheckCircle2 className="w-16 h-16 mx-auto mb-4 text-green-500" />
                    <p className="text-lg font-medium text-slate-900 mb-1">{fileName}</p>
                    <p className="text-sm text-green-600">{rawData.length} students loaded • Click to replace</p>
                  </>
                ) : (
                  <>
                    <Upload className="w-16 h-16 mx-auto mb-4 text-blue-500 group-hover:scale-110 transition-transform" />
                    <p className="text-lg font-medium text-slate-900 mb-2">
                      Drop your file here or click to browse
                    </p>
                    <p className="text-sm text-slate-500">
                      Supports CSV, XLSX, and XLS files
                    </p>
                  </>
                )}
              </div>
            </div>

            {/* Proceed Button */}
            <button
              onClick={proceedToConfigure}
              disabled={rawData.length === 0}
              className="w-full px-6 py-4 bg-gradient-to-r from-blue-600 to-indigo-600 text-white font-semibold rounded-xl hover:from-blue-700 hover:to-indigo-700 disabled:from-slate-300 disabled:to-slate-400 disabled:cursor-not-allowed transition-all shadow-lg shadow-blue-500/25"
            >
              Continue to Configuration →
            </button>
          </div>
        )}

        {/* ═══════════════════ Configure State ═══════════════════ */}
        {state === 'configure' && (
          <div className="max-w-4xl mx-auto space-y-6">
            <div className="bg-white rounded-2xl shadow-lg shadow-blue-100/50 border border-slate-200 p-8">
              <div className="flex items-center gap-3 mb-6">
                <Settings className="w-6 h-6 text-blue-600" />
                <h2 className="text-2xl font-bold text-slate-900">Configure Grading Parameters</h2>
              </div>

              <div className="mb-8 p-4 bg-blue-50 border border-blue-200 rounded-lg">
                <div className="flex items-start gap-3">
                  <AlertCircle className="w-5 h-5 text-blue-600 mt-0.5" />
                  <div>
                    <p className="text-sm font-medium text-blue-900 mb-1">File Loaded: {fileName}</p>
                    <p className="text-sm text-blue-700">{rawData.length} students • {numericColumns.length} numeric columns detected</p>
                  </div>
                </div>
              </div>

              {/* Name Column Selection */}
              <div className="mb-8">
                <label className="block text-sm font-semibold text-slate-900 mb-3">
                  Select Student Name Column
                </label>
                <select
                  value={nameColumn}
                  onChange={(e) => setNameColumn(e.target.value)}
                  className="w-full px-4 py-3 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-900"
                >
                  <option value="">-- Select Name Column --</option>
                  {allColumns.map(col => (
                    <option key={col} value={col}>{col}</option>
                  ))}
                </select>
              </div>

              {/* Roll Number Column Selection (conditional on team projects) */}
              {includeTeamProjects && (
                <div className="mb-8 animate-fadeIn">
                  <label className="block text-sm font-semibold text-slate-900 mb-3">
                    Select Roll Number Column
                  </label>
                  <select
                    value={rollNumberColumn}
                    onChange={(e) => setRollNumberColumn(e.target.value)}
                    className="w-full px-4 py-3 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-900"
                  >
                    <option value="">-- Select Roll Number Column --</option>
                    {allColumns.map(col => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                  <p className="text-xs text-slate-500 mt-1">
                    Roll numbers must end with numeric digits (e.g. 25BCS018). The last digits are used to match team members.
                  </p>
                </div>
              )}

              {/* Column Selection */}
              <div>
                <label className="block text-sm font-semibold text-slate-900 mb-3">
                  Select Columns to Grade
                </label>
                <div className="space-y-3">
                  {numericColumns.map(column => (
                    <div key={column} className="border border-slate-200 rounded-lg p-4 hover:border-blue-300 transition-colors">
                      <div className="flex items-center gap-3 mb-3">
                        <input
                          type="checkbox"
                          id={column}
                          checked={selectedColumns.has(column)}
                          onChange={() => handleColumnToggle(column)}
                          className="w-5 h-5 text-blue-600 rounded border-slate-300 focus:ring-2 focus:ring-blue-500"
                        />
                        <label htmlFor={column} className="font-medium text-slate-900 flex-1 cursor-pointer">
                          {column}
                        </label>
                      </div>
                      
                      {selectedColumns.has(column) && (
                        <div className="ml-8 grid grid-cols-2 gap-4 animate-fadeIn">
                          <div>
                            <label className="block text-xs font-medium text-slate-600 mb-1">
                              Max Marks
                            </label>
                            <input
                              type="number"
                              min="1"
                              value={columnConfigs.get(column)?.maxMarks ?? 100}
                              onChange={(e) => updateColumnConfig(column, 'maxMarks', e.target.value)}
                              onBlur={() => handleConfigBlur(column, 'maxMarks')}
                              className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-900"
                            />
                          </div>
                          <div>
                            <label className="block text-xs font-medium text-slate-600 mb-1">
                              Weightage (%)
                            </label>
                            <input
                              type="number"
                              min="0"
                              max="100"
                              step="0.1"
                              value={columnConfigs.get(column)?.weightage ?? 0}
                              onChange={(e) => updateColumnConfig(column, 'weightage', e.target.value)}
                              onBlur={() => handleConfigBlur(column, 'weightage')}
                              className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-900"
                            />
                          </div>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>

              <button
                onClick={processData}
                disabled={selectedColumns.size === 0 || !nameColumn || (includeTeamProjects && (!rollNumberColumn || !teamMarksColumn || hasTeamBlockingErrors || teamRawData.length === 0))}
                className="w-full mt-8 px-6 py-4 bg-gradient-to-r from-blue-600 to-indigo-600 text-white font-semibold rounded-lg hover:from-blue-700 hover:to-indigo-700 disabled:from-slate-300 disabled:to-slate-400 disabled:cursor-not-allowed transition-all shadow-lg shadow-blue-500/25"
              >
                Calculate Grades →
              </button>
            </div>

            {/* ─── Include Team Projects Toggle + File Upload ─── */}
            <div className="bg-white rounded-2xl shadow-lg shadow-purple-100/50 border border-slate-200 p-8">
              <div className="flex items-center gap-3 mb-2">
                <input
                  type="checkbox"
                  id="includeTeamProjects"
                  checked={includeTeamProjects}
                  onChange={(e) => setIncludeTeamProjects(e.target.checked)}
                  className="w-5 h-5 text-purple-600 rounded border-slate-300 focus:ring-2 focus:ring-purple-500"
                />
                <label htmlFor="includeTeamProjects" className="text-lg font-semibold text-slate-900 cursor-pointer flex items-center gap-2">
                  <Users className="w-5 h-5 text-purple-600" />
                  Include Team Projects?
                </label>
              </div>
              <p className="text-xs text-slate-500 ml-8 mb-4">Enable this to factor in team/group project scores alongside individual marks.</p>

              {includeTeamProjects && (
                <div className="ml-8 animate-fadeIn">
                  <label className="block text-sm font-semibold text-slate-900 mb-3">
                    Upload Team Projects File
                  </label>
                  <div
                    className={`border-2 border-dashed rounded-xl p-8 text-center hover:border-purple-500 hover:bg-purple-50/50 transition-all cursor-pointer group ${teamFileName ? 'border-green-400 bg-green-50/30' : 'border-purple-300'}`}
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={(e) => {
                      e.preventDefault();
                      const file = e.dataTransfer.files[0];
                      if (file) handleTeamFileUpload(file);
                    }}
                    onClick={() => {
                      const input = document.createElement('input');
                      input.type = 'file';
                      input.accept = '.csv,.xlsx,.xls';
                      input.onchange = (e) => {
                        const file = (e.target as HTMLInputElement).files?.[0];
                        if (file) handleTeamFileUpload(file);
                      };
                      input.click();
                    }}
                  >
                    {teamFileName ? (
                      <>
                        <CheckCircle2 className="w-12 h-12 mx-auto mb-3 text-green-500" />
                        <p className="text-base font-medium text-slate-900 mb-1">{teamFileName}</p>
                        <p className="text-sm text-green-600">{teamRawData.length} teams loaded • Click to replace</p>
                      </>
                    ) : (
                      <>
                        <Users className="w-12 h-12 mx-auto mb-3 text-purple-500 group-hover:scale-110 transition-transform" />
                        <p className="text-base font-medium text-slate-900 mb-1">
                          Drop team file here or click to browse
                        </p>
                        <p className="text-sm text-slate-500">CSV or Excel with team members and marks</p>
                      </>
                    )}
                  </div>
                </div>
              )}
            </div>

            {/* ─── Team Project Configuration Card ─── */}
            {includeTeamProjects && teamRawData.length > 0 && (
              <div className="bg-white rounded-2xl shadow-lg shadow-purple-100/50 border border-purple-200 p-8 animate-fadeIn">
                <div className="flex items-center gap-3 mb-6">
                  <Users className="w-6 h-6 text-purple-600" />
                  <h3 className="text-xl font-bold text-slate-900">Team Project Configuration</h3>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6">
                  {/* Team Marks Column */}
                  <div>
                    <label className="block text-sm font-semibold text-slate-900 mb-2">
                      Select Team Marks Column
                    </label>
                    <select
                      value={teamMarksColumn}
                      onChange={(e) => setTeamMarksColumn(e.target.value)}
                      className="w-full px-4 py-3 border border-slate-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500 text-slate-900"
                    >
                      <option value="">-- Select Marks Column --</option>
                      {teamColumns.map(col => (
                        <option key={col} value={col}>{col}</option>
                      ))}
                    </select>
                  </div>

                  {/* Max Team Marks */}
                  <div>
                    <label className="block text-sm font-semibold text-slate-900 mb-2">
                      Max Team Marks
                    </label>
                    <input
                      type="number"
                      min="1"
                      value={maxTeamMarks}
                      onChange={(e) => setMaxTeamMarks(Number(e.target.value) || 100)}
                      className="w-full px-4 py-3 border border-slate-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500 text-slate-900"
                    />
                  </div>

                  {/* Team Weightage */}
                  <div>
                    <label className="block text-sm font-semibold text-slate-900 mb-2">
                      Team Weightage (%)
                    </label>
                    <input
                      type="number"
                      min="0"
                      max="100"
                      value={teamWeightage}
                      onChange={(e) => {
                        const val = Math.min(100, Math.max(0, Number(e.target.value) || 0));
                        setTeamWeightage(val);
                      }}
                      className="w-full px-4 py-3 border border-slate-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500 text-slate-900"
                    />
                    <p className="text-xs text-slate-500 mt-1">
                      Individual: {100 - teamWeightage}%
                    </p>
                  </div>
                </div>

                {/* Weightage Display */}
                <div className="p-4 bg-purple-50 border border-purple-200 rounded-lg mb-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-sm font-medium text-purple-900">Team Project Weightage</p>
                      <p className="text-2xl font-bold text-purple-700">{teamWeightage}%</p>
                    </div>
                    <div className="text-right">
                      <p className="text-sm font-medium text-blue-900">Individual Component Weightage</p>
                      <p className="text-2xl font-bold text-blue-700">{100 - teamWeightage}%</p>
                    </div>
                  </div>
                </div>

                {/* Team Matching Summary */}
                {teamMatchSummary && (
                  <div className={`p-4 border rounded-lg ${teamValidationErrors.length > 0 ? 'bg-amber-50 border-amber-200' : 'bg-green-50 border-green-200'}`}>
                    <div className="flex items-start gap-3">
                      {teamValidationErrors.length > 0 ? (
                        <AlertTriangle className="w-5 h-5 text-amber-600 mt-0.5 flex-shrink-0" />
                      ) : (
                        <CheckCircle2 className="w-5 h-5 text-green-600 mt-0.5 flex-shrink-0" />
                      )}
                      <div className="flex-1">
                        <p className="text-sm font-medium text-slate-900 mb-2">Team Matching Summary</p>
                        <div className="grid grid-cols-3 gap-4 mb-2">
                          <div className="text-center p-2 bg-white/60 rounded">
                            <p className="text-lg font-bold text-green-700">{teamMatchSummary.studentsInTeams}</p>
                            <p className="text-xs text-slate-600">Students in Teams</p>
                          </div>
                          <div className="text-center p-2 bg-white/60 rounded">
                            <p className="text-lg font-bold text-purple-700">{teamMatchSummary.totalTeams}</p>
                            <p className="text-xs text-slate-600">Total Teams</p>
                          </div>
                          <div className="text-center p-2 bg-white/60 rounded">
                            <p className="text-lg font-bold text-amber-700">{teamMatchSummary.studentsWithoutTeams}</p>
                            <p className="text-xs text-slate-600">Without Teams</p>
                          </div>
                        </div>
                        {teamValidationErrors.map((err, i) => (
                          <p key={i} className="text-sm text-amber-800 mt-1">⚠ {err}</p>
                        ))}
                        {teamMatchSummary.duplicateMembers.length > 0 && (
                          <p className="text-sm text-red-700 font-medium mt-2">
                            ❌ Cannot proceed: Fix duplicate member assignments before calculating grades.
                          </p>
                        )}
                      </div>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ═══════════════════ Dashboard State ═══════════════════ */}
        {state === 'dashboard' && (
          <div className="space-y-6">
            {/* Statistics Cards */}
            <div className={`grid grid-cols-1 gap-6 ${teamEnabledInResults ? 'md:grid-cols-5' : 'md:grid-cols-4'}`}>
              <div className="bg-white rounded-xl shadow-lg shadow-blue-100/50 border border-slate-200 p-6">
                <p className="text-sm font-medium text-slate-600 mb-1">Total Students</p>
                <p className="text-3xl font-bold text-slate-900">{processedData.length}</p>
              </div>
              <div className="bg-white rounded-xl shadow-lg shadow-blue-100/50 border border-slate-200 p-6">
                <p className="text-sm font-medium text-slate-600 mb-1">Mean Score (μ)</p>
                <p className="text-3xl font-bold text-blue-600">{mean.toFixed(2)}</p>
              </div>
              <div className="bg-white rounded-xl shadow-lg shadow-blue-100/50 border border-slate-200 p-6">
                <p className="text-sm font-medium text-slate-600 mb-1">Std Dev (σ)</p>
                <p className="text-3xl font-bold text-indigo-600">{stdDev.toFixed(2)}</p>
              </div>
              {teamDashboardStats && (
                <div className="bg-white rounded-xl shadow-lg shadow-purple-100/50 border border-purple-200 p-6">
                  <p className="text-sm font-medium text-purple-600 mb-1">Students in Teams</p>
                  <p className="text-3xl font-bold text-purple-700">
                    {teamDashboardStats.inTeams}<span className="text-lg text-slate-400">/{teamDashboardStats.total}</span>
                  </p>
                </div>
              )}
              <div className="bg-white rounded-xl shadow-lg shadow-blue-100/50 border border-slate-200 p-6">
                <button
                  onClick={openExportModal}
                  className="flex items-center justify-center gap-2 w-full h-full px-4 py-3 bg-gradient-to-r from-green-600 to-emerald-600 text-white font-semibold rounded-lg hover:from-green-700 hover:to-emerald-700 transition-all shadow-lg shadow-green-500/25"
                >
                  <Download className="w-5 h-5" />
                  Export Results
                </button>
              </div>
            </div>

            {/* Export Column Selection Modal */}
            {showExportModal && (
              <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-50 flex items-center justify-center p-4" onClick={() => setShowExportModal(false)}>
                <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-lg max-h-[80vh] flex flex-col animate-fadeIn" onClick={(e) => e.stopPropagation()}>
                  <div className="p-6 border-b border-slate-200">
                    <div className="flex items-center justify-between">
                      <h3 className="text-lg font-bold text-slate-900">Select Export Columns</h3>
                      <button onClick={() => setShowExportModal(false)} className="text-slate-400 hover:text-slate-600 transition-colors">
                        <X className="w-5 h-5" />
                      </button>
                    </div>
                    <p className="text-sm text-slate-500 mt-1">Choose which columns to include in the exported file.</p>
                    <div className="flex gap-3 mt-3">
                      <button
                        onClick={() => setExportSelectedColumns(new Set(availableExportColumns))}
                        className="text-xs px-3 py-1.5 bg-blue-50 text-blue-700 rounded-md hover:bg-blue-100 transition-colors font-medium"
                      >
                        Select All
                      </button>
                      <button
                        onClick={() => setExportSelectedColumns(new Set())}
                        className="text-xs px-3 py-1.5 bg-slate-50 text-slate-700 rounded-md hover:bg-slate-100 transition-colors font-medium"
                      >
                        Deselect All
                      </button>
                    </div>
                  </div>
                  <div className="p-6 overflow-y-auto flex-1">
                    <div className="space-y-2">
                      {availableExportColumns.map(col => (
                        <label key={col} className="flex items-center gap-3 p-2 rounded-lg hover:bg-slate-50 cursor-pointer transition-colors">
                          <input
                            type="checkbox"
                            checked={exportSelectedColumns.has(col)}
                            onChange={() => toggleExportColumn(col)}
                            className="w-4 h-4 text-blue-600 rounded border-slate-300 focus:ring-2 focus:ring-blue-500"
                          />
                          <span className="text-sm text-slate-800 font-medium">{col}</span>
                        </label>
                      ))}
                    </div>
                  </div>
                  <div className="p-6 border-t border-slate-200 flex gap-3 justify-end">
                    <button
                      onClick={() => setShowExportModal(false)}
                      className="px-5 py-2.5 text-sm font-medium text-slate-700 bg-slate-100 rounded-lg hover:bg-slate-200 transition-colors"
                    >
                      Cancel
                    </button>
                    <button
                      onClick={exportToExcel}
                      disabled={exportSelectedColumns.size === 0}
                      className="flex items-center gap-2 px-5 py-2.5 text-sm font-medium text-white bg-gradient-to-r from-green-600 to-emerald-600 rounded-lg hover:from-green-700 hover:to-emerald-700 disabled:from-slate-300 disabled:to-slate-400 disabled:cursor-not-allowed transition-all shadow-md shadow-green-500/20"
                    >
                      <Download className="w-4 h-4" />
                      Download ({exportSelectedColumns.size} columns)
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Stacked Histogram Chart */}
            <div className="bg-white rounded-2xl shadow-lg shadow-blue-100/50 border border-slate-200 p-8">
              <div className="flex items-center justify-between mb-2">
                <h3 className="text-xl font-bold text-slate-900">Score Distribution Histogram</h3>
                <div className="flex items-center gap-1 text-xs text-slate-500">
                  <span className="inline-block w-3 h-3 rounded-sm" style={{ backgroundColor: GRADE_COLORS['A'] }} />
                  <span>Excellent</span>
                  <span className="mx-1">→</span>
                  <span className="inline-block w-3 h-3 rounded-sm" style={{ backgroundColor: GRADE_COLORS['F'] }} />
                  <span>Fail</span>
                </div>
              </div>
              <p className="text-sm text-slate-500 mb-6">Each bar shows how many students fall within a 5-point score range, color-coded by grade.</p>
              <div className="overflow-x-auto">
                {(() => {
                  // Build histogram data grouped into 5-point score bins
                  const GRADE_KEYS = Object.keys(GRADE_COLORS) as Array<keyof typeof GRADE_COLORS>;
                  const BIN_SIZE = 5;

                  // Determine range
                  const scores = processedData.map(s => s.finalScore);
                  const minScore = Math.floor(Math.min(...scores) / BIN_SIZE) * BIN_SIZE;
                  const maxScore = Math.ceil(Math.max(...scores) / BIN_SIZE) * BIN_SIZE;

                  // Build bins
                  const bins = new Map<number, Record<string, number | string>>();
                  for (let b = minScore; b < maxScore; b += BIN_SIZE) {
                    const entry: Record<string, number | string> = { binStart: b, label: `${b}–${b + BIN_SIZE}` };
                    GRADE_KEYS.forEach(g => { entry[g] = 0; });
                    bins.set(b, entry);
                  }

                  processedData.forEach(student => {
                    const bin = Math.floor(student.finalScore / BIN_SIZE) * BIN_SIZE;
                    const clampedBin = Math.max(minScore, Math.min(bin, maxScore - BIN_SIZE));
                    const entry = bins.get(clampedBin);
                    if (entry) {
                      const grade = student.grade as keyof typeof GRADE_COLORS;
                      if (typeof entry[grade] === 'number') {
                        (entry[grade] as number)++;
                      }
                    }
                  });

                  const histogramData = Array.from(bins.values()).sort((a, b) => (a.binStart as number) - (b.binStart as number));

                  return (
                    <BarChart
                      width={1000}
                      height={600}
                      data={histogramData}
                      margin={{ top: 30, right: 30, left: 30, bottom: 50 }}
                      barCategoryGap="12%"
                    >
                      <defs>
                        {GRADE_KEYS.map(grade => (
                          <linearGradient key={`grad-${grade}`} id={`gradient-${grade}`} x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor={GRADE_COLORS[grade]} stopOpacity={1} />
                            <stop offset="100%" stopColor={GRADE_COLORS[grade]} stopOpacity={0.7} />
                          </linearGradient>
                        ))}
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                      <XAxis
                        dataKey="label"
                        tick={{ fill: '#334155', fontSize: 13, fontWeight: 600 }}
                        tickLine={false}
                        axisLine={{ stroke: '#cbd5e1' }}
                        label={{ value: 'Score Range', position: 'insideBottom', offset: -15, fill: '#334155', fontSize: 14, fontWeight: 700 }}
                      />
                      <YAxis
                        allowDecimals={false}
                        tick={{ fill: '#334155', fontSize: 13, fontWeight: 500 }}
                        tickLine={false}
                        axisLine={{ stroke: '#cbd5e1' }}
                        label={{ value: 'Number of Students', angle: -90, position: 'insideLeft', offset: -5, fill: '#334155', fontSize: 14, fontWeight: 700 }}
                      />
                      <Tooltip content={<HistogramChartTooltip />} />
                      <Legend
                        formatter={(value: string) => <span style={{ color: '#334155', fontWeight: 600, fontSize: 13 }}>{value} ({GRADE_LABELS[value] || value})</span>}
                        iconType="square"
                        iconSize={14}
                        wrapperStyle={{ paddingTop: 16 }}
                      />
                      {GRADE_KEYS.map((grade, i) => (
                        <Bar
                          key={grade}
                          dataKey={grade}
                          name={grade}
                          stackId="a"
                          fill={`url(#gradient-${grade})`}
                          stroke={GRADE_COLORS[grade]}
                          strokeWidth={1}
                          radius={i === GRADE_KEYS.length - 1 ? [6, 6, 0, 0] : [0, 0, 0, 0]}
                        />
                      ))}
                    </BarChart>
                  );
                })()}
              </div>
            </div>

            {/* Cutoff Sliders */}
            <div className="bg-white rounded-2xl shadow-lg shadow-blue-100/50 border border-slate-200 p-8">
              <h3 className="text-xl font-bold text-slate-900 mb-6">Adjust Grade Cutoffs</h3>
              <div className="grid grid-cols-1 gap-4">
                {(Object.keys(cutoffs) as Array<keyof GradeCutoffs>).map(grade => (
                  <div key={grade} className="space-y-2">
                    <div className="flex items-center justify-between">
                      <label className="text-sm font-semibold text-slate-900">
                        Grade {grade}
                      </label>
                      <span className="text-sm font-mono font-medium text-slate-600 bg-slate-100 px-3 py-1 rounded-md">
                        {cutoffs[grade].toFixed(2)}
                      </span>
                    </div>
                    <input
                      type="range"
                      min="0"
                      max="100"
                      step="0.1"
                      value={cutoffs[grade]}
                      onChange={(e) => setCutoffs({ ...cutoffs, [grade]: Number(e.target.value) })}
                      className="w-full h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-blue-600"
                      style={{
                        background: `linear-gradient(to right, ${GRADE_COLORS[grade === 'D' ? 'D' : grade]} 0%, ${GRADE_COLORS[grade === 'D' ? 'D' : grade]} ${(cutoffs[grade] / 100) * 100}%, #e2e8f0 ${(cutoffs[grade] / 100) * 100}%, #e2e8f0 100%)`
                      }}
                    />
                  </div>
                ))}
              </div>
            </div>

            {/* Grade Distribution Table */}
            <div className="bg-white rounded-2xl shadow-lg shadow-blue-100/50 border border-slate-200 p-8">
              <h3 className="text-xl font-bold text-slate-900 mb-6">Grade Distribution</h3>
              <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-8 gap-4">
                {Object.entries(GRADE_COLORS).map(([grade, color]) => {
                  const count = processedData.filter(s => s.grade === grade).length;
                  const percentage = ((count / processedData.length) * 100).toFixed(1);
                  return (
                    <div key={grade} className="text-center p-4 rounded-lg border border-slate-200 hover:shadow-md transition-shadow">
                      <div className="w-12 h-12 mx-auto mb-2 rounded-full flex items-center justify-center text-slate-900 font-bold text-lg" style={{ backgroundColor: color, opacity: 0.85 }}>
                        {grade}
                      </div>
                      <p className="text-2xl font-bold text-slate-900">{count}</p>
                      <p className="text-xs text-slate-600">{percentage}%</p>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </main>

      <style jsx global>{`
        @keyframes fadeIn {
          from { opacity: 0; transform: translateY(-10px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-fadeIn {
          animation: fadeIn 0.3s ease-out;
        }
      `}</style>
    </div>
  );
}
