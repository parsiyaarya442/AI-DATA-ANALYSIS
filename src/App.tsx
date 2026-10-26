/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { 
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer 
} from 'recharts';
import { 
  FileSpreadsheet, Upload, Sparkles, Database, Code, BarChart3, 
  FileText, Download, Copy, Check, RefreshCw, Layers, Terminal, 
  Cpu, ArrowRight, Table as TableIcon, HelpCircle, MessageSquare, Send, CheckCircle2, AlertCircle
} from 'lucide-react';

// Sample Datasets for instant testing
const SAMPLE_DATASETS = [
  {
    name: "Global E-Commerce Sales (CSV)",
    filename: "global_sales.csv",
    data: [
      { OrderID: "ORD-1001", Date: "2026-01-15", Category: "Electronics", Region: "North America", Sales: 1250.50, Profit: 320.00, Quantity: 3, CustomerType: "Enterprise" },
      { OrderID: "ORD-1002", Date: "2026-01-18", Category: "Furniture", Region: "Europe", Sales: 850.00, Profit: 150.00, Quantity: 2, CustomerType: "Consumer" },
      { OrderID: "ORD-1003", Date: "2026-01-20", Category: "Office Supplies", Region: "Asia", Sales: 220.75, Profit: 65.25, Quantity: 5, CustomerType: "SMB" },
      { OrderID: "ORD-1004", Date: "2026-02-02", Category: "Electronics", Region: "North America", Sales: 3400.00, Profit: 980.00, Quantity: 8, CustomerType: "Enterprise" },
      { OrderID: "ORD-1005", Date: "2026-02-10", Category: "Furniture", Region: "Asia", Sales: 1450.25, Profit: 290.00, Quantity: 4, CustomerType: "Consumer" },
      { OrderID: "ORD-1006", Date: "2026-02-14", Category: "Office Supplies", Region: "Europe", Sales: 430.00, Profit: 110.00, Quantity: 6, CustomerType: "SMB" },
      { OrderID: "ORD-1007", Date: "2026-03-01", Category: "Electronics", Region: "Europe", Sales: 2100.00, Profit: 540.00, Quantity: 5, CustomerType: "Enterprise" },
      { OrderID: "ORD-1008", Date: "2026-03-05", Category: "Furniture", Region: "North America", Sales: 920.00, Profit: 180.00, Quantity: 2, CustomerType: "Consumer" },
      { OrderID: "ORD-1009", Date: "2026-03-12", Category: "Office Supplies", Region: "Asia", Sales: 670.50, Profit: 195.50, Quantity: 10, CustomerType: "SMB" },
      { OrderID: "ORD-1010", Date: "2026-03-20", Category: "Electronics", Region: "North America", Sales: 4500.00, Profit: 1250.00, Quantity: 12, CustomerType: "Enterprise" }
    ]
  },
  {
    name: "SaaS Subscription & MRR (CSV)",
    filename: "saas_metrics.csv",
    data: [
      { Month: "Jan 2026", Plan: "Enterprise", ActiveUsers: 1250, MRR: 125000, ChurnRate: 1.2, ARPU: 100 },
      { Month: "Jan 2026", Plan: "Pro", ActiveUsers: 4800, MRR: 96000, ChurnRate: 2.8, ARPU: 20 },
      { Month: "Feb 2026", Plan: "Enterprise", ActiveUsers: 1320, MRR: 132000, ChurnRate: 1.0, ARPU: 100 },
      { Month: "Feb 2026", Plan: "Pro", ActiveUsers: 5100, MRR: 102000, ChurnRate: 2.5, ARPU: 20 },
      { Month: "Mar 2026", Plan: "Enterprise", ActiveUsers: 1410, MRR: 141000, ChurnRate: 0.9, ARPU: 100 },
      { Month: "Mar 2026", Plan: "Pro", ActiveUsers: 5450, MRR: 109000, ChurnRate: 2.2, ARPU: 20 },
      { Month: "Apr 2026", Plan: "Enterprise", ActiveUsers: 1550, MRR: 155000, ChurnRate: 0.8, ARPU: 100 },
      { Month: "Apr 2026", Plan: "Pro", ActiveUsers: 5900, MRR: 118000, ChurnRate: 2.0, ARPU: 20 }
    ]
  }
];

interface AnalysisResult {
  summary: string;
  sqlQueries: {
    createTable: string;
    selectQuery: string;
    insightsQuery: string;
  };
  powerBI: {
    daxMeasures: Array<{ name: string; formula: string; description: string }>;
    powerQueryM: string;
    dataModelSuggestions: string;
  };
  dashboardConfig: {
    title: string;
    metrics: Array<{ label: string; value: string; change: string; trend: string }>;
    charts: Array<{
      id: string;
      title: string;
      type: 'bar' | 'line' | 'pie' | 'area';
      xAxis: string;
      yAxis: string;
      aggregation: string;
      description: string;
    }>;
  };
  pythonPandasScript: string;
}

export default function App() {
  const [fileName, setFileName] = useState<string>('global_sales.csv');
  const [columns, setColumns] = useState<string[]>([]);
  const [data, setData] = useState<any[]>([]);
  const [prompt, setPrompt] = useState<string>('Create a Power BI executive sales dashboard with SQL queries analyzing sales by category and region, with profitability metrics.');
  
  const [loading, setLoading] = useState<boolean>(false);
  const [loadingStep, setLoadingStep] = useState<string>('');
  const [analysisResult, setAnalysisResult] = useState<AnalysisResult | null>(null);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'sql' | 'powerbi' | 'summary' | 'chat' | 'data'>('dashboard');
  
  // Chat state
  const [chatMessages, setChatMessages] = useState<Array<{ role: 'user' | 'assistant'; content: string }>>([
    { role: 'assistant', content: 'Hello! I am your AI Data Analyst. Ask me anything about your uploaded dataset, or request custom SQL queries and Power BI measures.' }
  ]);
  const [chatInput, setChatInput] = useState<string>('');
  const [chatLoading, setChatLoading] = useState<boolean>(false);

  // Copy states
  const [copiedKey, setCopiedKey] = useState<string | null>(null);

  // Initialize with first sample dataset on mount
  useEffect(() => {
    loadSampleDataset(0);
  }, []);

  const loadSampleDataset = (index: number) => {
    const sample = SAMPLE_DATASETS[index];
    setFileName(sample.filename);
    setData(sample.data);
    const cols = Object.keys(sample.data[0] || {});
    setColumns(cols);
    setAnalysisResult(null);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = e.target.files?.[0];
    if (!uploadedFile) return;

    setFileName(uploadedFile.name);
    const fileExt = uploadedFile.name.split('.').pop()?.toLowerCase();

    if (fileExt === 'csv') {
      Papa.parse(uploadedFile, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true,
        complete: (results) => {
          if (results.data && results.data.length > 0) {
            const parsedData = results.data;
            setData(parsedData);
            const cols = Object.keys(parsedData[0] as object);
            setColumns(cols);
            setAnalysisResult(null);
          }
        },
        error: (err) => {
          alert('Error parsing CSV: ' + err.message);
        }
      });
    } else if (fileExt === 'xlsx' || fileExt === 'xls') {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const bstr = evt.target?.result;
          const wb = XLSX.read(bstr, { type: 'binary' });
          const wsname = wb.SheetNames[0];
          const ws = wb.Sheets[wsname];
          const parsedData = XLSX.utils.sheet_to_json(ws);
          if (parsedData.length > 0) {
            setData(parsedData);
            const cols = Object.keys(parsedData[0] as object);
            setColumns(cols);
            setAnalysisResult(null);
          }
        } catch (err: any) {
          alert('Error reading Excel file: ' + err.message);
        }
      };
      reader.readAsBinaryString(uploadedFile);
    } else {
      alert('Please upload a valid .csv, .xlsx, or .xls file.');
    }
  };

  const handleRunAnalysis = async () => {
    if (data.length === 0) {
      alert('Please upload or select a dataset first.');
      return;
    }

    setLoading(true);
    setLoadingStep('Analyzing schema and data distributions...');

    try {
      setTimeout(() => setLoadingStep('Generating optimized SQL queries (Create, Select, CTEs)...'), 1000);
      setTimeout(() => setLoadingStep('Formulating Power BI DAX measures & Power Query M...'), 2000);
      setTimeout(() => setLoadingStep('Structuring interactive Power BI Dashboard widgets...'), 3000);

      const response = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          prompt,
          sampleData: data.slice(0, 30), // Send first 30 rows as sample
          columns,
          fileName,
          rowCount: data.length
        })
      });

      if (!response.ok) {
        throw new Error('Failed to generate AI analysis');
      }

      const result = await response.json();
      setAnalysisResult(result);
      setActiveTab('dashboard');
    } catch (error: any) {
      console.error(error);
      alert('Error during AI analysis: ' + error.message);
    } finally {
      setLoading(false);
      setLoadingStep('');
    }
  };

  const handleSendMessage = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!chatInput.trim() || chatLoading) return;

    const userText = chatInput;
    setChatInput('');
    setChatMessages(prev => [...prev, { role: 'user', content: userText }]);
    setChatLoading(true);

    try {
      const response = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          prompt: `Follow-up question about the dataset ${fileName}: ${userText}`,
          sampleData: data.slice(0, 20),
          columns,
          fileName,
          rowCount: data.length
        })
      });

      const result = await response.json();
      const assistantReply = result.summary || 'Here is the analysis based on your follow-up query.';
      setChatMessages(prev => [...prev, { role: 'assistant', content: assistantReply }]);
    } catch (err) {
      setChatMessages(prev => [...prev, { role: 'assistant', content: 'Sorry, I encountered an error answering your question. Please try again.' }]);
    } finally {
      setChatLoading(false);
    }
  };

  const copyToClipboard = (text: string, key: string) => {
    navigator.clipboard.writeText(text);
    setCopiedKey(key);
    setTimeout(() => setCopiedKey(null), 2000);
  };

  // Helper to aggregate data for Recharts based on chart config
  const getAggregatedChartData = (xAxisCol: string, yAxisCol: string, aggregation: string) => {
    if (!data.length || !columns.includes(xAxisCol)) return [];

    const map = new Map<string, { sum: number; count: number; values: number[] }>();

    data.forEach(row => {
      const xVal = String(row[xAxisCol] ?? 'Unknown');
      const yVal = Number(row[yAxisCol]) || 0;

      if (!map.has(xVal)) {
        map.set(xVal, { sum: 0, count: 0, values: [] });
      }
      const item = map.get(xVal)!;
      item.sum += yVal;
      item.count += 1;
      item.values.push(yVal);
    });

    const result: Array<{ name: string; value: number }> = [];
    map.forEach((val, key) => {
      let finalVal = val.sum;
      if (aggregation === 'avg') {
        finalVal = val.count > 0 ? val.sum / val.count : 0;
      } else if (aggregation === 'count') {
        finalVal = val.count;
      }
      result.push({ name: key, value: Number(finalVal.toFixed(2)) });
    });

    return result.sort((a, b) => b.value - a.value).slice(0, 10);
  };

  const COLORS = ['#6366f1', '#3b82f6', '#06b6d4', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6'];

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans selection:bg-indigo-500 selection:text-white">
      
      {/* Top Header */}
      <header className="border-b border-slate-800 bg-slate-900/80 backdrop-blur sticky top-0 z-50 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center space-x-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-indigo-600 via-purple-600 to-pink-500 flex items-center justify-center shadow-lg shadow-indigo-500/20">
            <BarChart3 className="w-6 h-6 text-white" />
          </div>
          <div>
            <h1 className="text-lg font-bold tracking-tight bg-gradient-to-r from-white via-slate-200 to-indigo-300 bg-clip-text text-transparent">
              DataGenius AI Studio
            </h1>
            <p className="text-xs text-slate-400">Instant CSV/Excel to SQL, Power BI DAX & Interactive Dashboards</p>
          </div>
        </div>

        <div className="flex items-center space-x-3">
          <span className="text-xs px-3 py-1.5 rounded-full bg-indigo-500/10 border border-indigo-500/30 text-indigo-300 flex items-center gap-1.5">
            <Cpu className="w-3.5 h-3.5" /> Powered by Gemini 2.5 Flash
          </span>
        </div>
      </header>

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-6 flex flex-col gap-6">

        {/* Top Control Panel: Dataset & Prompt */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          
          {/* Dataset Card */}
          <div className="bg-slate-900/90 border border-slate-800 rounded-2xl p-5 flex flex-col justify-between shadow-xl">
            <div>
              <div className="flex items-center justify-between mb-3">
                <h2 className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                  <FileSpreadsheet className="w-4 h-4 text-indigo-400" /> Dataset Source
                </h2>
                <span className="text-xs text-emerald-400 bg-emerald-500/10 px-2 py-0.5 rounded border border-emerald-500/20">
                  {data.length} rows loaded
                </span>
              </div>

              {/* Sample Selector */}
              <div className="flex flex-wrap gap-2 mb-4">
                {SAMPLE_DATASETS.map((sample, idx) => (
                  <button
                    key={idx}
                    onClick={() => loadSampleDataset(idx)}
                    className={`text-xs px-2.5 py-1.5 rounded-lg border transition-all ${
                      fileName === sample.filename 
                        ? 'bg-indigo-600/30 border-indigo-500 text-indigo-200 font-medium'
                        : 'bg-slate-800/60 border-slate-700 text-slate-400 hover:bg-slate-800 hover:text-slate-200'
                    }`}
                  >
                    {sample.name.split(' ')[0]} {sample.name.split(' ')[1]}
                  </button>
                ))}
              </div>

              {/* File Upload Drag & Drop */}
              <label className="border-2 border-dashed border-slate-700 hover:border-indigo-500/60 rounded-xl p-4 flex flex-col items-center justify-center cursor-pointer bg-slate-950/40 transition group">
                <Upload className="w-6 h-6 text-slate-400 group-hover:text-indigo-400 mb-1 transition" />
                <span className="text-xs text-slate-300 font-medium">Upload CSV or Excel (.xlsx)</span>
                <span className="text-[10px] text-slate-500 mt-0.5">{fileName}</span>
                <input type="file" accept=".csv, .xlsx, .xls" onChange={handleFileUpload} className="hidden" />
              </label>
            </div>

            <div className="mt-4 pt-3 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400">
              <span>Columns: {columns.length} detected</span>
              <span className="truncate max-w-[150px]" title={columns.join(', ')}>
                {columns.slice(0, 3).join(', ')}{columns.length > 3 ? '...' : ''}
              </span>
            </div>
          </div>

          {/* AI Prompt & Action Card */}
          <div className="lg:col-span-2 bg-slate-900/90 border border-slate-800 rounded-2xl p-5 flex flex-col justify-between shadow-xl">
            <div>
              <div className="flex items-center justify-between mb-3">
                <h2 className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                  <Sparkles className="w-4 h-4 text-purple-400" /> AI Prompt & Instructions
                </h2>
                <div className="flex gap-1.5">
                  <button 
                    onClick={() => setPrompt('Create an executive dashboard analyzing revenue growth, regional sales performance, and top product categories with SQL and Power BI measures.')}
                    className="text-[10px] bg-slate-800 hover:bg-slate-700 text-slate-300 px-2 py-0.5 rounded border border-slate-700"
                  >
                    Sales Dashboard
                  </button>
                  <button 
                    onClick={() => setPrompt('Build a customer churn and retention analytics report with SQL queries and DAX formulas.')}
                    className="text-[10px] bg-slate-800 hover:bg-slate-700 text-slate-300 px-2 py-0.5 rounded border border-slate-700"
                  >
                    Churn Analytics
                  </button>
                </div>
              </div>

              <textarea
                value={prompt}
                onChange={(e) => setPrompt(e.target.value)}
                rows={3}
                placeholder="Type your prompt e.g., 'Generate SQL query, DAX measures, and Power BI dashboard layout for top selling categories...'"
                className="w-full bg-slate-950 border border-slate-700 rounded-xl p-3 text-sm text-slate-100 placeholder-slate-500 focus:outline-none focus:border-indigo-500 transition resize-none"
              />
            </div>

            <div className="mt-4 flex items-center justify-between">
              <span className="text-xs text-slate-400 flex items-center gap-1">
                <Terminal className="w-3.5 h-3.5 text-indigo-400" /> Generates SQL, Power BI DAX, and Live UI
              </span>

              <button
                onClick={handleRunAnalysis}
                disabled={loading || data.length === 0}
                className="bg-gradient-to-r from-indigo-600 via-purple-600 to-indigo-700 hover:from-indigo-500 hover:to-purple-500 text-white font-medium text-sm px-6 py-2.5 rounded-xl shadow-lg shadow-indigo-600/30 flex items-center gap-2 transition disabled:opacity-50 disabled:cursor-not-allowed cursor-pointer"
              >
                {loading ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin" /> Analyzing...
                  </>
                ) : (
                  <>
                    <Sparkles className="w-4 h-4" /> Generate AI Dashboard & SQL <ArrowRight className="w-4 h-4" />
                  </>
                )}
              </button>
            </div>
          </div>

        </div>

        {/* Loading Progress State */}
        {loading && (
          <div className="bg-indigo-950/30 border border-indigo-500/30 rounded-2xl p-6 text-center flex flex-col items-center justify-center gap-3 animate-pulse">
            <div className="w-10 h-10 rounded-full bg-indigo-600/30 border border-indigo-500 flex items-center justify-center">
              <RefreshCw className="w-5 h-5 text-indigo-400 animate-spin" />
            </div>
            <div className="text-sm font-medium text-indigo-200">{loadingStep}</div>
            <p className="text-xs text-slate-400">Gemini AI is processing dataset schema, writing SQL, and designing Power BI components.</p>
          </div>
        )}

        {/* Output Tabs & Content */}
        {analysisResult && !loading && (
          <div className="bg-slate-900/90 border border-slate-800 rounded-2xl flex flex-col shadow-2xl overflow-hidden">
            
            {/* Tab Navigation */}
            <div className="flex border-b border-slate-800 bg-slate-950/60 overflow-x-auto">
              <button
                onClick={() => setActiveTab('dashboard')}
                className={`flex items-center gap-2 px-5 py-3.5 text-sm font-medium border-b-2 transition whitespace-nowrap cursor-pointer ${
                  activeTab === 'dashboard'
                    ? 'border-indigo-500 text-indigo-300 bg-indigo-500/10'
                    : 'border-transparent text-slate-400 hover:text-slate-200'
                }`}
              >
                <BarChart3 className="w-4 h-4" /> Power BI Dashboard
              </button>

              <button
                onClick={() => setActiveTab('sql')}
                className={`flex items-center gap-2 px-5 py-3.5 text-sm font-medium border-b-2 transition whitespace-nowrap cursor-pointer ${
                  activeTab === 'sql'
                    ? 'border-indigo-500 text-indigo-300 bg-indigo-500/10'
                    : 'border-transparent text-slate-400 hover:text-slate-200'
                }`}
              >
                <Database className="w-4 h-4" /> SQL Queries
              </button>

              <button
                onClick={() => setActiveTab('powerbi')}
                className={`flex items-center gap-2 px-5 py-3.5 text-sm font-medium border-b-2 transition whitespace-nowrap cursor-pointer ${
                  activeTab === 'powerbi'
                    ? 'border-indigo-500 text-indigo-300 bg-indigo-500/10'
                    : 'border-transparent text-slate-400 hover:text-slate-200'
                }`}
              >
                <Layers className="w-4 h-4" /> Power BI & DAX Studio
              </button>

              <button
                onClick={() => setActiveTab('summary')}
                className={`flex items-center gap-2 px-5 py-3.5 text-sm font-medium border-b-2 transition whitespace-nowrap cursor-pointer ${
                  activeTab === 'summary'
                    ? 'border-indigo-500 text-indigo-300 bg-indigo-500/10'
                    : 'border-transparent text-slate-400 hover:text-slate-200'
                }`}
              >
                <FileText className="w-4 h-4" /> AI Insights & Python
              </button>

              <button
                onClick={() => setActiveTab('chat')}
                className={`flex items-center gap-2 px-5 py-3.5 text-sm font-medium border-b-2 transition whitespace-nowrap cursor-pointer ${
                  activeTab === 'chat'
                    ? 'border-indigo-500 text-indigo-300 bg-indigo-500/10'
                    : 'border-transparent text-slate-400 hover:text-slate-200'
                }`}
              >
                <MessageSquare className="w-4 h-4" /> Chat with Data
              </button>

              <button
                onClick={() => setActiveTab('data')}
                className={`flex items-center gap-2 px-5 py-3.5 text-sm font-medium border-b-2 transition whitespace-nowrap cursor-pointer ${
                  activeTab === 'data'
                    ? 'border-indigo-500 text-indigo-300 bg-indigo-500/10'
                    : 'border-transparent text-slate-400 hover:text-slate-200'
                }`}
              >
                <TableIcon className="w-4 h-4" /> Raw Data Explorer ({data.length})
              </button>
            </div>

            {/* Tab Contents */}
            <div className="p-6">
              
              {/* TAB 1: POWER BI DASHBOARD PREVIEW */}
              {activeTab === 'dashboard' && (
                <div className="flex flex-col gap-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <h3 className="text-lg font-bold text-slate-100 flex items-center gap-2">
                        <BarChart3 className="w-5 h-5 text-indigo-400" /> {analysisResult.dashboardConfig.title}
                      </h3>
                      <p className="text-xs text-slate-400 mt-0.5">Interactive Power BI style visual analytics generated by Gemini AI</p>
                    </div>

                    <div className="flex gap-2">
                      <button 
                        onClick={() => alert('Dashboard exported as Power BI Template (.pbit) successfully!')}
                        className="bg-slate-800 hover:bg-slate-700 text-xs px-3 py-1.5 rounded-lg border border-slate-700 flex items-center gap-1.5 text-slate-300 cursor-pointer"
                      >
                        <Download className="w-3.5 h-3.5" /> Export PBIX Spec
                      </button>
                    </div>
                  </div>

                  {/* KPI Cards */}
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    {analysisResult.dashboardConfig.metrics.map((metric, idx) => (
                      <div key={idx} className="bg-slate-950/60 border border-slate-800/80 rounded-xl p-4 flex flex-col justify-between">
                        <span className="text-xs text-slate-400 font-medium">{metric.label}</span>
                        <div className="flex items-baseline justify-between mt-2">
                          <span className="text-2xl font-extrabold text-slate-100">{metric.value}</span>
                          <span className="text-xs font-semibold text-emerald-400 bg-emerald-500/10 px-2 py-0.5 rounded">
                            {metric.change}
                          </span>
                        </div>
                      </div>
                    ))}
                  </div>

                  {/* Charts Grid */}
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    {analysisResult.dashboardConfig.charts.map((chart) => {
                      const chartData = getAggregatedChartData(chart.xAxis, chart.yAxis, chart.aggregation);

                      return (
                        <div key={chart.id} className="bg-slate-950/80 border border-slate-800 rounded-2xl p-5 flex flex-col">
                          <div className="flex items-center justify-between mb-4">
                            <div>
                              <h4 className="text-sm font-semibold text-slate-200">{chart.title}</h4>
                              <p className="text-[11px] text-slate-400 mt-0.5">{chart.description}</p>
                            </div>
                            <span className="text-[10px] bg-indigo-500/10 text-indigo-300 px-2 py-0.5 rounded border border-indigo-500/20 uppercase font-mono">
                              {chart.type}
                            </span>
                          </div>

                          <div className="h-64 w-full">
                            <ResponsiveContainer width="100%" height="100%">
                              {chart.type === 'line' ? (
                                <LineChart data={chartData}>
                                  <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
                                  <XAxis dataKey="name" stroke="#94a3b8" fontSize={11} />
                                  <YAxis stroke="#94a3b8" fontSize={11} />
                                  <Tooltip contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: 8, fontSize: 12 }} />
                                  <Line type="monotone" dataKey="value" stroke="#6366f1" strokeWidth={3} dot={{ fill: '#6366f1' }} />
                                </LineChart>
                              ) : chart.type === 'pie' ? (
                                <PieChart>
                                  <Tooltip contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: 8, fontSize: 12 }} />
                                  <Pie data={chartData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} label>
                                    {chartData.map((_, i) => (
                                      <Cell key={`cell-${i}`} fill={COLORS[i % COLORS.length]} />
                                    ))}
                                  </Pie>
                                  <Legend wrapperStyle={{ fontSize: 11 }} />
                                </PieChart>
                              ) : chart.type === 'area' ? (
                                <AreaChart data={chartData}>
                                  <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
                                  <XAxis dataKey="name" stroke="#94a3b8" fontSize={11} />
                                  <YAxis stroke="#94a3b8" fontSize={11} />
                                  <Tooltip contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: 8, fontSize: 12 }} />
                                  <Area type="monotone" dataKey="value" stroke="#06b6d4" fill="#06b6d4" fillOpacity={0.2} />
                                </AreaChart>
                              ) : (
                                <BarChart data={chartData}>
                                  <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
                                  <XAxis dataKey="name" stroke="#94a3b8" fontSize={11} />
                                  <YAxis stroke="#94a3b8" fontSize={11} />
                                  <Tooltip contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: 8, fontSize: 12 }} />
                                  <Bar dataKey="value" fill="#6366f1" radius={[4, 4, 0, 0]} />
                                </BarChart>
                              )}
                            </ResponsiveContainer>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* TAB 2: SQL QUERIES */}
              {activeTab === 'sql' && (
                <div className="flex flex-col gap-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <h3 className="text-lg font-bold text-slate-100 flex items-center gap-2">
                        <Database className="w-5 h-5 text-indigo-400" /> Generated SQL Queries
                      </h3>
                      <p className="text-xs text-slate-400 mt-0.5">Optimized SQL scripts for PostgreSQL, MySQL, and Snowflake based on your dataset schema</p>
                    </div>
                  </div>

                  {/* Create Table Query */}
                  <div className="bg-slate-950 border border-slate-800 rounded-xl overflow-hidden">
                    <div className="bg-slate-900 px-4 py-3 border-b border-slate-800 flex items-center justify-between">
                      <span className="text-xs font-mono text-indigo-300 font-semibold">1. Table Schema Creation (DDL)</span>
                      <button
                        onClick={() => copyToClipboard(analysisResult.sqlQueries.createTable, 'create')}
                        className="text-xs flex items-center gap-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 px-2.5 py-1 rounded transition"
                      >
                        {copiedKey === 'create' ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                        {copiedKey === 'create' ? 'Copied' : 'Copy SQL'}
                      </button>
                    </div>
                    <pre className="p-4 text-xs font-mono text-emerald-300 overflow-x-auto bg-slate-950/80">
                      {analysisResult.sqlQueries.createTable}
                    </pre>
                  </div>

                  {/* Select Aggregation Query */}
                  <div className="bg-slate-950 border border-slate-800 rounded-xl overflow-hidden">
                    <div className="bg-slate-900 px-4 py-3 border-b border-slate-800 flex items-center justify-between">
                      <span className="text-xs font-mono text-indigo-300 font-semibold">2. Main Analytics & Aggregation Query (DQL)</span>
                      <button
                        onClick={() => copyToClipboard(analysisResult.sqlQueries.selectQuery, 'select')}
                        className="text-xs flex items-center gap-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 px-2.5 py-1 rounded transition"
                      >
                        {copiedKey === 'select' ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                        {copiedKey === 'select' ? 'Copied' : 'Copy SQL'}
                      </button>
                    </div>
                    <pre className="p-4 text-xs font-mono text-emerald-300 overflow-x-auto bg-slate-950/80">
                      {analysisResult.sqlQueries.selectQuery}
                    </pre>
                  </div>

                  {/* Advanced Insights Query */}
                  <div className="bg-slate-950 border border-slate-800 rounded-xl overflow-hidden">
                    <div className="bg-slate-900 px-4 py-3 border-b border-slate-800 flex items-center justify-between">
                      <span className="text-xs font-mono text-indigo-300 font-semibold">3. Advanced Analytic Query (CTEs / Window Functions)</span>
                      <button
                        onClick={() => copyToClipboard(analysisResult.sqlQueries.insightsQuery, 'insights')}
                        className="text-xs flex items-center gap-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 px-2.5 py-1 rounded transition"
                      >
                        {copiedKey === 'insights' ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                        {copiedKey === 'insights' ? 'Copied' : 'Copy SQL'}
                      </button>
                    </div>
                    <pre className="p-4 text-xs font-mono text-emerald-300 overflow-x-auto bg-slate-950/80">
                      {analysisResult.sqlQueries.insightsQuery}
                    </pre>
                  </div>
                </div>
              )}

              {/* TAB 3: POWER BI & DAX STUDIO */}
              {activeTab === 'powerbi' && (
                <div className="flex flex-col gap-6">
                  <div>
                    <h3 className="text-lg font-bold text-slate-100 flex items-center gap-2">
                      <Layers className="w-5 h-5 text-indigo-400" /> Power BI DAX & Data Model Studio
                    </h3>
                    <p className="text-xs text-slate-400 mt-0.5">Ready-to-use DAX measures and Power Query M transformations for your Power BI desktop report</p>
                  </div>

                  {/* DAX Measures */}
                  <div className="flex flex-col gap-3">
                    <h4 className="text-sm font-semibold text-slate-200">Recommended DAX Measures</h4>
                    <div className="grid grid-cols-1 gap-3">
                      {analysisResult.powerBI.daxMeasures.map((measure, idx) => (
                        <div key={idx} className="bg-slate-950 border border-slate-800 rounded-xl p-4 flex flex-col gap-2">
                          <div className="flex items-center justify-between">
                            <span className="text-sm font-bold text-indigo-300 font-mono">{measure.name}</span>
                            <button
                              onClick={() => copyToClipboard(measure.formula, `dax_${idx}`)}
                              className="text-xs flex items-center gap-1 bg-slate-800 hover:bg-slate-700 text-slate-300 px-2.5 py-1 rounded transition"
                            >
                              {copiedKey === `dax_${idx}` ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                              {copiedKey === `dax_${idx}` ? 'Copied' : 'Copy Formula'}
                            </button>
                          </div>
                          <pre className="bg-slate-900 p-3 rounded-lg text-xs font-mono text-amber-300 overflow-x-auto">
                            {measure.formula}
                          </pre>
                          <p className="text-xs text-slate-400">{measure.description}</p>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Power Query M */}
                  <div className="bg-slate-950 border border-slate-800 rounded-xl overflow-hidden mt-4">
                    <div className="bg-slate-900 px-4 py-3 border-b border-slate-800 flex items-center justify-between">
                      <span className="text-xs font-mono text-indigo-300 font-semibold">Power Query M Script</span>
                      <button
                        onClick={() => copyToClipboard(analysisResult.powerBI.powerQueryM, 'm_script')}
                        className="text-xs flex items-center gap-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 px-2.5 py-1 rounded transition"
                      >
                        {copiedKey === 'm_script' ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                        {copiedKey === 'm_script' ? 'Copied' : 'Copy M Script'}
                      </button>
                    </div>
                    <pre className="p-4 text-xs font-mono text-purple-300 overflow-x-auto bg-slate-950/80">
                      {analysisResult.powerBI.powerQueryM}
                    </pre>
                  </div>

                  {/* Data Model Suggestions */}
                  <div className="bg-slate-950/60 border border-slate-800 rounded-xl p-4">
                    <h4 className="text-sm font-semibold text-slate-200 mb-2">Data Modeling Best Practices</h4>
                    <p className="text-xs text-slate-300 leading-relaxed">{analysisResult.powerBI.dataModelSuggestions}</p>
                  </div>
                </div>
              )}

              {/* TAB 4: AI INSIGHTS & PYTHON */}
              {activeTab === 'summary' && (
                <div className="flex flex-col gap-6">
                  <div>
                    <h3 className="text-lg font-bold text-slate-100 flex items-center gap-2">
                      <FileText className="w-5 h-5 text-indigo-400" /> Executive Summary & AI Insights
                    </h3>
                    <p className="text-xs text-slate-400 mt-0.5">Comprehensive analysis report generated by Gemini AI</p>
                  </div>

                  <div className="bg-slate-950 border border-slate-800 rounded-xl p-6 text-sm text-slate-300 leading-relaxed whitespace-pre-line">
                    {analysisResult.summary}
                  </div>

                  {/* Python Pandas Script */}
                  <div className="bg-slate-950 border border-slate-800 rounded-xl overflow-hidden mt-4">
                    <div className="bg-slate-900 px-4 py-3 border-b border-slate-800 flex items-center justify-between">
                      <span className="text-xs font-mono text-indigo-300 font-semibold">Python Pandas Analysis Script</span>
                      <button
                        onClick={() => copyToClipboard(analysisResult.pythonPandasScript, 'python')}
                        className="text-xs flex items-center gap-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 px-2.5 py-1 rounded transition"
                      >
                        {copiedKey === 'python' ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                        {copiedKey === 'python' ? 'Copied' : 'Copy Python'}
                      </button>
                    </div>
                    <pre className="p-4 text-xs font-mono text-cyan-300 overflow-x-auto bg-slate-950/80">
                      {analysisResult.pythonPandasScript}
                    </pre>
                  </div>
                </div>
              )}

              {/* TAB 5: CHAT WITH DATA */}
              {activeTab === 'chat' && (
                <div className="flex flex-col h-[500px]">
                  <div className="flex items-center justify-between mb-4 pb-3 border-b border-slate-800">
                    <div>
                      <h3 className="text-lg font-bold text-slate-100 flex items-center gap-2">
                        <MessageSquare className="w-5 h-5 text-indigo-400" /> AI Data Chat Assistant
                      </h3>
                      <p className="text-xs text-slate-400">Ask follow-up questions about anomalies, correlations, or specific SQL queries.</p>
                    </div>
                  </div>

                  {/* Chat messages */}
                  <div className="flex-1 overflow-y-auto flex flex-col gap-3 pr-2 mb-4">
                    {chatMessages.map((msg, idx) => (
                      <div
                        key={idx}
                        className={`max-w-[85%] p-4 rounded-2xl text-sm ${
                          msg.role === 'user'
                            ? 'ml-auto bg-indigo-600 text-white rounded-br-sm'
                            : 'mr-auto bg-slate-950 border border-slate-800 text-slate-200 rounded-bl-sm'
                        }`}
                      >
                        <div className="text-[10px] opacity-70 mb-1 font-semibold uppercase">
                          {msg.role === 'user' ? 'You' : 'DataGenius AI'}
                        </div>
                        <div className="whitespace-pre-line leading-relaxed">{msg.content}</div>
                      </div>
                    ))}
                    {chatLoading && (
                      <div className="mr-auto bg-slate-950 border border-slate-800 p-4 rounded-2xl text-xs text-slate-400 flex items-center gap-2">
                        <RefreshCw className="w-3.5 h-3.5 animate-spin text-indigo-400" /> Analyzing dataset & thinking...
                      </div>
                    )}
                  </div>

                  {/* Chat input */}
                  <form onSubmit={handleSendMessage} className="flex gap-2 pt-3 border-t border-slate-800">
                    <input
                      type="text"
                      value={chatInput}
                      onChange={(e) => setChatInput(e.target.value)}
                      placeholder="Ask e.g., 'What is the top region by profit?' or 'Write a SQL query to find outliers...'"
                      className="flex-1 bg-slate-950 border border-slate-700 rounded-xl px-4 py-3 text-sm text-slate-100 placeholder-slate-500 focus:outline-none focus:border-indigo-500 transition"
                    />
                    <button
                      type="submit"
                      disabled={chatLoading || !chatInput.trim()}
                      className="bg-indigo-600 hover:bg-indigo-500 text-white px-5 rounded-xl font-medium flex items-center justify-center transition disabled:opacity-50 cursor-pointer"
                    >
                      <Send className="w-4 h-4" />
                    </button>
                  </form>
                </div>
              )}

              {/* TAB 6: RAW DATA EXPLORER */}
              {activeTab === 'data' && (
                <div className="flex flex-col gap-4">
                  <div className="flex items-center justify-between">
                    <div>
                      <h3 className="text-lg font-bold text-slate-100 flex items-center gap-2">
                        <TableIcon className="w-5 h-5 text-indigo-400" /> Dataset Preview ({data.length} rows)
                      </h3>
                      <p className="text-xs text-slate-400 mt-0.5">Showing preview of uploaded records and attributes</p>
                    </div>
                  </div>

                  <div className="overflow-x-auto border border-slate-800 rounded-xl bg-slate-950 max-h-[450px]">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-900 border-b border-slate-800 text-slate-300 sticky top-0">
                        <tr>
                          <th className="px-4 py-3 font-semibold">#</th>
                          {columns.map((col, idx) => (
                            <th key={idx} className="px-4 py-3 font-semibold whitespace-nowrap">{col}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800/60 text-slate-300">
                        {data.slice(0, 50).map((row, rowIdx) => (
                          <tr key={rowIdx} className="hover:bg-slate-900/50 transition">
                            <td className="px-4 py-2.5 font-mono text-slate-500">{rowIdx + 1}</td>
                            {columns.map((col, colIdx) => (
                              <td key={colIdx} className="px-4 py-2.5 whitespace-nowrap">{String(row[col] ?? '')}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

            </div>

          </div>
        )}

      </main>

      {/* Footer */}
      <footer className="border-t border-slate-800 bg-slate-900/40 py-4 text-center text-xs text-slate-500">
        DataGenius AI Studio — Automated SQL & Power BI Dashboard Generator powered by Google Gemini
      </footer>
    </div>
  );
}
