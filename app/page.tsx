"use client";

import React, { useState, useRef, ChangeEvent, DragEvent } from "react";

interface Step1Stats {
  brinkRows: number;
  tableauRows: number;
  totalRows: number;
  reconciledStores: number;
  latestDateStr: string;
}

export default function Home() {
  // Step 1: Franchise Sales state
  const [brinkFile, setBrinkFile] = useState<File | null>(null);
  const [tableauFile, setTableauFile] = useState<File | null>(null);
  const [isDraggingBrink, setIsDraggingBrink] = useState(false);
  const [isDraggingTableau, setIsDraggingTableau] = useState(false);
  const [isProcessingStep1, setIsProcessingStep1] = useState(false);
  const [step1Done, setStep1Done] = useState(false);
  const [step1Stats, setStep1Stats] = useState<Step1Stats | null>(null);
  const [step1DownloadUrl, setStep1DownloadUrl] = useState<string | null>(null);
  const [step1FileName, setStep1FileName] = useState<string>("Franchise Sales (Updated).xlsx");
  const [reconciledTableauRows, setReconciledTableauRows] = useState<any[] | null>(null);

  // Step 2: Royalty Calculator state
  const [royaltyFile, setRoyaltyFile] = useState<File | null>(null);
  const [isDraggingRoyalty, setIsDraggingRoyalty] = useState(false);
  const [isProcessingStep2, setIsProcessingStep2] = useState(false);
  const [step2Done, setStep2Done] = useState(false);
  const [step2DownloadUrl, setStep2DownloadUrl] = useState<string | null>(null);
  const [step2FileName, setStep2FileName] = useState<string | null>(null);
  const [step2RecordsCount, setStep2RecordsCount] = useState<number | null>(null);

  // General error message
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const brinkInputRef = useRef<HTMLInputElement>(null);
  const tableauInputRef = useRef<HTMLInputElement>(null);
  const royaltyInputRef = useRef<HTMLInputElement>(null);

  const validateExcelFile = (file: File): boolean => {
    const lower = file.name.toLowerCase();
    return lower.endsWith(".xlsx") || lower.endsWith(".xls");
  };

  const handleBrinkChange = (e: ChangeEvent<HTMLInputElement>) => {
    setErrorMessage(null);
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (validateExcelFile(file)) setBrinkFile(file);
      else setErrorMessage("Please select a valid Excel file (.xlsx or .xls) for Brink.");
    }
  };

  const handleTableauChange = (e: ChangeEvent<HTMLInputElement>) => {
    setErrorMessage(null);
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (validateExcelFile(file)) setTableauFile(file);
      else setErrorMessage("Please select a valid Excel file (.xlsx or .xls) for Tableau.");
    }
  };

  const handleRoyaltyChange = (e: ChangeEvent<HTMLInputElement>) => {
    setErrorMessage(null);
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (validateExcelFile(file)) setRoyaltyFile(file);
      else setErrorMessage("Please select a valid Excel file (.xlsx or .xls) for Royalty Calculator.");
    }
  };

  const handleBrinkDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDraggingBrink(false);
    setErrorMessage(null);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (validateExcelFile(file)) setBrinkFile(file);
      else setErrorMessage("Please select a valid Excel file for Brink.");
    }
  };

  const handleTableauDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDraggingTableau(false);
    setErrorMessage(null);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (validateExcelFile(file)) setTableauFile(file);
      else setErrorMessage("Please select a valid Excel file for Tableau.");
    }
  };

  const handleRoyaltyDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDraggingRoyalty(false);
    setErrorMessage(null);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (validateExcelFile(file)) setRoyaltyFile(file);
      else setErrorMessage("Please select a valid Excel file for Royalty Calculator.");
    }
  };

  // Step 1: Process Brink + Tableau -> Franchise Sales
  const handleProcessStep1 = async () => {
    if (!brinkFile || !tableauFile) {
      setErrorMessage("Both Brink Sales Summary and Tableau Data files are required for Step 1.");
      return;
    }

    setIsProcessingStep1(true);
    setErrorMessage(null);

    try {
      const formData = new FormData();
      formData.append("brinkFile", brinkFile);
      formData.append("tableauFile", tableauFile);

      const response = await fetch("/api/update-sales", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let errText = "Failed to process files in Step 1.";
        try {
          const json = await response.json();
          if (json.error) errText = json.error;
        } catch {
          errText = await response.text();
        }
        throw new Error(errText);
      }

      const resData = await response.json();
      if (!resData.success) {
        throw new Error(resData.error || "Failed to process Franchise Sales file.");
      }

      // Trigger automatic download of Franchise Sales (Updated).xlsx
      const byteCharacters = atob(resData.file.data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: resData.file.mimeType });
      const downloadUrl = window.URL.createObjectURL(blob);

      const downloadLink = document.createElement("a");
      downloadLink.href = downloadUrl;
      downloadLink.download = resData.file.fileName;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);

      setStep1DownloadUrl(downloadUrl);
      setStep1FileName(resData.file.fileName);
      setStep1Stats(resData.stats);
      setReconciledTableauRows(resData.tableauRows);
      setStep1Done(true);
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : "Error processing Step 1.");
    } finally {
      setIsProcessingStep1(false);
    }
  };

  // Step 2: Upload Royalty Calculator + Ingest Tableau rows -> Download Royalty Calculator
  const handleProcessStep2 = async () => {
    if (!royaltyFile) {
      setErrorMessage("Please select a Franchise Royalty Calculator file to upload.");
      return;
    }
    if (!reconciledTableauRows || reconciledTableauRows.length === 0) {
      setErrorMessage("Please complete Step 1 first to prepare the Tableau sales data.");
      return;
    }

    setIsProcessingStep2(true);
    setErrorMessage(null);

    try {
      const formData = new FormData();
      formData.append("royaltyFile", royaltyFile);
      formData.append("tableauRows", JSON.stringify(reconciledTableauRows));

      const response = await fetch("/api/update-royalty", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let errText = "Failed to process Franchise Royalty Calculator.";
        try {
          const json = await response.json();
          if (json.error) errText = json.error;
        } catch {
          errText = await response.text();
        }
        throw new Error(errText);
      }

      const resData = await response.json();
      if (!resData.success) {
        throw new Error(resData.error || "Failed to process Royalty Calculator.");
      }

      // Trigger automatic download of Franchise Royalty Calculator <date>.xlsx
      const byteCharacters = atob(resData.file.data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: resData.file.mimeType });
      const downloadUrl = window.URL.createObjectURL(blob);

      const downloadLink = document.createElement("a");
      downloadLink.href = downloadUrl;
      downloadLink.download = resData.file.fileName;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);

      setStep2DownloadUrl(downloadUrl);
      setStep2FileName(resData.file.fileName);
      setStep2RecordsCount(resData.stats?.recordsCount || reconciledTableauRows.length);
      setStep2Done(true);
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : "Error processing Step 2.");
    } finally {
      setIsProcessingStep2(false);
    }
  };

  const handleResetAll = () => {
    setBrinkFile(null);
    setTableauFile(null);
    setRoyaltyFile(null);
    setStep1Done(false);
    setStep2Done(false);
    setStep1Stats(null);
    setStep1DownloadUrl(null);
    setStep2DownloadUrl(null);
    setStep2FileName(null);
    setReconciledTableauRows(null);
    setErrorMessage(null);
    if (brinkInputRef.current) brinkInputRef.current.value = "";
    if (tableauInputRef.current) tableauInputRef.current.value = "";
    if (royaltyInputRef.current) royaltyInputRef.current.value = "";
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  return (
    <div className="h-screen max-h-screen overflow-hidden flex flex-col bg-zinc-100 p-3 sm:p-4 text-zinc-900 select-none">
      <div className="max-w-4xl w-full mx-auto flex flex-col h-full justify-between gap-2.5">
        
        {/* Header Bar */}
        <header className="bg-white rounded-xl px-4 py-2.5 shadow-2xs border border-zinc-200 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2.5">
            <span className="flex h-8 w-8 items-center justify-center rounded-lg bg-emerald-50 text-emerald-600 font-semibold text-base border border-emerald-100 shadow-2xs">
              📊
            </span>
            <div>
              <h1 className="text-sm sm:text-base font-bold tracking-tight text-zinc-900 leading-tight">
                Franchise Sales &amp; Royalty Hub
              </h1>
              <p className="text-[11px] text-zinc-400 font-mono">
                Two-Step Automated Workflow
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            {/* Step Indicators */}
            <div className="flex items-center gap-1.5 text-xs font-semibold">
              <span
                className={`px-2.5 py-1 rounded-full border transition-all ${
                  step1Done
                    ? "bg-emerald-50 text-emerald-700 border-emerald-200 font-bold"
                    : "bg-emerald-600 text-white border-emerald-600 shadow-xs"
                }`}
              >
                {step1Done ? "✓ Step 1 Done" : "1. Sales"}
              </span>
              <span className="text-zinc-300">→</span>
              <span
                className={`px-2.5 py-1 rounded-full border transition-all ${
                  step2Done
                    ? "bg-emerald-50 text-emerald-700 border-emerald-200 font-bold"
                    : step1Done
                    ? "bg-blue-600 text-white border-blue-600 shadow-xs animate-pulse"
                    : "bg-zinc-100 text-zinc-400 border-zinc-200"
                }`}
              >
                {step2Done ? "✓ Step 2 Done" : "2. Royalty"}
              </span>
            </div>

            {(brinkFile || tableauFile || royaltyFile || step1Done) && (
              <button
                onClick={handleResetAll}
                className="ml-2 text-xs px-2.5 py-1 rounded-lg border border-zinc-200 text-zinc-500 hover:text-zinc-800 hover:bg-zinc-50 transition-colors cursor-pointer"
              >
                Reset
              </button>
            )}
          </div>
        </header>

        {/* Global Error Banner */}
        {errorMessage && (
          <div className="rounded-lg border border-red-200 bg-red-50 px-3 py-1.5 text-xs text-red-600 shrink-0 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span>⚠️</span>
              <span className="font-semibold">{errorMessage}</span>
            </div>
            <button
              onClick={() => setErrorMessage(null)}
              className="text-red-400 hover:text-red-700 font-bold text-sm leading-none ml-2"
            >
              ×
            </button>
          </div>
        )}

        {/* STEP 1 CARD: Franchise Sales (Both Brink & Tableau Required) */}
        <section className={`bg-white rounded-xl shadow-2xs border p-3.5 sm:p-4 flex flex-col justify-between flex-1 transition-all ${
          step1Done ? "border-emerald-200 bg-emerald-50/15" : "border-zinc-200"
        }`}>
          <div>
            <div className="flex items-center justify-between pb-2 border-b border-zinc-100 mb-2.5">
              <div className="flex items-center gap-2">
                <span className="flex h-5 w-5 items-center justify-center rounded-full bg-emerald-600 text-white text-xs font-bold">
                  1
                </span>
                <span className="font-bold text-xs sm:text-sm text-zinc-900">
                  Step 1: Upload Brink &amp; Tableau Files
                </span>
                <span className="text-[10px] uppercase font-bold tracking-wider bg-amber-50 text-amber-700 border border-amber-200 px-1.5 py-0.5 rounded">
                  Both Required
                </span>
              </div>

              {step1Done && step1Stats && (
                <div className="flex items-center gap-2 text-[11px] font-medium text-emerald-700">
                  <span className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded font-bold">
                    ✓ Reconciled {step1Stats.reconciledStores} Stores
                  </span>
                  {step1DownloadUrl && (
                    <a
                      href={step1DownloadUrl}
                      download={step1FileName}
                      className="inline-flex items-center gap-1 font-semibold text-emerald-700 hover:text-emerald-900 underline"
                    >
                      📥 Re-download Sales
                    </a>
                  )}
                </div>
              )}
            </div>

            {/* Dual Dropzones */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
              {/* 1. Brink Upload */}
              <div className="flex flex-col">
                <div className="flex items-center justify-between mb-1">
                  <span className="text-[11px] font-bold text-zinc-700 uppercase tracking-wider flex items-center gap-1">
                    <span>📊</span> Brink Sales Summary
                  </span>
                  <span className="text-[10px] bg-zinc-100 text-zinc-500 px-1.5 py-0.2 rounded font-mono">
                    Sheet #3
                  </span>
                </div>
                <input
                  ref={brinkInputRef}
                  type="file"
                  accept=".xlsx, .xls"
                  onChange={handleBrinkChange}
                  className="hidden"
                  id="brink-input"
                />
                <div
                  onDragOver={(e) => { e.preventDefault(); setIsDraggingBrink(true); }}
                  onDragLeave={() => setIsDraggingBrink(false)}
                  onDrop={handleBrinkDrop}
                  onClick={() => brinkInputRef.current?.click()}
                  className={`h-20 sm:h-22 border-2 border-dashed rounded-lg p-2 flex items-center justify-center text-center cursor-pointer transition-all ${
                    isDraggingBrink
                      ? "border-emerald-500 bg-emerald-50/60"
                      : brinkFile
                      ? "border-emerald-300 bg-emerald-50/30"
                      : "border-zinc-300 hover:border-zinc-400 bg-zinc-50/70 hover:bg-zinc-50"
                  }`}
                >
                  <div className="flex items-center gap-2.5">
                    <div className={`w-8 h-8 rounded-full flex items-center justify-center text-sm shrink-0 ${
                      brinkFile ? "bg-emerald-600 text-white" : "bg-zinc-200 text-zinc-500"
                    }`}>
                      {brinkFile ? "✓" : "📁"}
                    </div>
                    <div className="text-left overflow-hidden">
                      {brinkFile ? (
                        <>
                          <p className="font-semibold text-zinc-900 text-xs truncate max-w-[180px] sm:max-w-[220px]">
                            {brinkFile.name}
                          </p>
                          <p className="text-[10px] text-zinc-400">
                            {formatFileSize(brinkFile.size)} • Click to change
                          </p>
                        </>
                      ) : (
                        <>
                          <p className="font-semibold text-zinc-800 text-xs">
                            Select Brink File
                          </p>
                          <p className="text-[10px] text-zinc-400">
                            Drop .xlsx / .xls
                          </p>
                        </>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              {/* 2. Tableau Upload */}
              <div className="flex flex-col">
                <div className="flex items-center justify-between mb-1">
                  <span className="text-[11px] font-bold text-zinc-700 uppercase tracking-wider flex items-center gap-1">
                    <span>📈</span> Tableau Data
                  </span>
                  <span className="text-[10px] bg-emerald-50 text-emerald-700 font-semibold px-1.5 py-0.2 rounded font-mono">
                    Sheet #6
                  </span>
                </div>
                <input
                  ref={tableauInputRef}
                  type="file"
                  accept=".xlsx, .xls"
                  onChange={handleTableauChange}
                  className="hidden"
                  id="tableau-input"
                />
                <div
                  onDragOver={(e) => { e.preventDefault(); setIsDraggingTableau(true); }}
                  onDragLeave={() => setIsDraggingTableau(false)}
                  onDrop={handleTableauDrop}
                  onClick={() => tableauInputRef.current?.click()}
                  className={`h-20 sm:h-22 border-2 border-dashed rounded-lg p-2 flex items-center justify-center text-center cursor-pointer transition-all ${
                    isDraggingTableau
                      ? "border-emerald-500 bg-emerald-50/60"
                      : tableauFile
                      ? "border-emerald-300 bg-emerald-50/30"
                      : "border-zinc-300 hover:border-zinc-400 bg-zinc-50/70 hover:bg-zinc-50"
                  }`}
                >
                  <div className="flex items-center gap-2.5">
                    <div className={`w-8 h-8 rounded-full flex items-center justify-center text-sm shrink-0 ${
                      tableauFile ? "bg-emerald-600 text-white" : "bg-zinc-200 text-zinc-500"
                    }`}>
                      {tableauFile ? "✓" : "📈"}
                    </div>
                    <div className="text-left overflow-hidden">
                      {tableauFile ? (
                        <>
                          <p className="font-semibold text-zinc-900 text-xs truncate max-w-[180px] sm:max-w-[220px]">
                            {tableauFile.name}
                          </p>
                          <p className="text-[10px] text-zinc-400">
                            {formatFileSize(tableauFile.size)} • Click to change
                          </p>
                        </>
                      ) : (
                        <>
                          <p className="font-semibold text-zinc-800 text-xs">
                            Select Tableau File
                          </p>
                          <p className="text-[10px] text-zinc-400">
                            Drop .xlsx / .xls
                          </p>
                        </>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Step 1 Action Button */}
          <div className="mt-2.5">
            <button
              onClick={handleProcessStep1}
              disabled={!brinkFile || !tableauFile || isProcessingStep1}
              className={`w-full py-2 px-3 rounded-lg font-semibold text-xs sm:text-sm flex items-center justify-center gap-2 transition-all shadow-xs ${
                !brinkFile || !tableauFile || isProcessingStep1
                  ? "bg-zinc-200 text-zinc-400 cursor-not-allowed"
                  : step1Done
                  ? "bg-emerald-700 hover:bg-emerald-800 text-white cursor-pointer"
                  : "bg-emerald-600 hover:bg-emerald-700 text-white cursor-pointer active:scale-[0.99]"
              }`}
            >
              {isProcessingStep1 ? (
                <>
                  <svg className="animate-spin h-3.5 w-3.5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"></path>
                  </svg>
                  Processing &amp; Reconciling Sales...
                </>
              ) : step1Done ? (
                "✓ Step 1 Complete (Re-run & Download Franchise Sales)"
              ) : (
                "⚡ Step 1: Process & Download Franchise Sales"
              )}
            </button>
          </div>
        </section>

        {/* STEP 2 CARD: Franchise Royalty Calculator (Unlocked after Step 1) */}
        <section className={`bg-white rounded-xl shadow-2xs border p-3.5 sm:p-4 flex flex-col justify-between flex-1 transition-all ${
          !step1Done
            ? "border-zinc-200 bg-zinc-50/50 opacity-60"
            : step2Done
            ? "border-blue-200 bg-blue-50/15"
            : "border-blue-300 bg-white ring-2 ring-blue-500/10"
        }`}>
          <div>
            <div className="flex items-center justify-between pb-2 border-b border-zinc-100 mb-2.5">
              <div className="flex items-center gap-2">
                <span className={`flex h-5 w-5 items-center justify-center rounded-full text-xs font-bold ${
                  step1Done ? "bg-blue-600 text-white" : "bg-zinc-300 text-zinc-600"
                }`}>
                  2
                </span>
                <span className="font-bold text-xs sm:text-sm text-zinc-900">
                  Step 2: Upload &amp; Update Franchise Royalty Calculator
                </span>
              </div>

              {!step1Done ? (
                <span className="text-[10px] font-semibold text-zinc-400 bg-zinc-100 px-2 py-0.5 rounded">
                  🔒 Complete Step 1 First
                </span>
              ) : step2Done && step2FileName ? (
                <div className="flex items-center gap-2 text-[11px] font-medium text-blue-700">
                  <span className="bg-blue-100 text-blue-800 px-2 py-0.5 rounded font-bold">
                    ✓ Injected {step2RecordsCount} Records
                  </span>
                  {step2DownloadUrl && (
                    <a
                      href={step2DownloadUrl}
                      download={step2FileName}
                      className="inline-flex items-center gap-1 font-semibold text-blue-700 hover:text-blue-900 underline truncate max-w-[200px]"
                    >
                      📥 Re-download ({step2FileName.replace("Franchise Royalty Calculator ", "")})
                    </a>
                  )}
                </div>
              ) : (
                <span className="text-[10px] font-bold text-blue-700 bg-blue-50 border border-blue-200 px-2 py-0.5 rounded">
                  Ready for Calculator File
                </span>
              )}
            </div>

            {/* Royalty Calculator Dropzone */}
            <div className="flex flex-col">
              <div className="flex items-center justify-between mb-1">
                <span className="text-[11px] font-bold text-zinc-700 uppercase tracking-wider flex items-center gap-1">
                  <span>📑</span> Franchise Royalty Calculator File
                </span>
                <span className="text-[10px] bg-blue-50 text-blue-700 font-semibold px-1.5 py-0.2 rounded font-mono">
                  Calculations Sheet (5 Columns)
                </span>
              </div>

              <input
                ref={royaltyInputRef}
                type="file"
                accept=".xlsx, .xls"
                disabled={!step1Done}
                onChange={handleRoyaltyChange}
                className="hidden"
                id="royalty-input"
              />
              <div
                onDragOver={(e) => {
                  if (step1Done) { e.preventDefault(); setIsDraggingRoyalty(true); }
                }}
                onDragLeave={() => setIsDraggingRoyalty(false)}
                onDrop={step1Done ? handleRoyaltyDrop : undefined}
                onClick={() => { if (step1Done) royaltyInputRef.current?.click(); }}
                className={`h-20 sm:h-22 border-2 border-dashed rounded-lg p-2 flex items-center justify-center text-center transition-all ${
                  !step1Done
                    ? "border-zinc-200 bg-zinc-100/50 cursor-not-allowed"
                    : isDraggingRoyalty
                    ? "border-blue-500 bg-blue-50/60 cursor-pointer"
                    : royaltyFile
                    ? "border-blue-300 bg-blue-50/30 cursor-pointer"
                    : "border-zinc-300 hover:border-blue-400 bg-zinc-50/70 hover:bg-zinc-50 cursor-pointer"
                }`}
              >
                <div className="flex items-center gap-2.5">
                  <div className={`w-8 h-8 rounded-full flex items-center justify-center text-sm shrink-0 ${
                    !step1Done
                      ? "bg-zinc-200 text-zinc-400"
                      : royaltyFile
                      ? "bg-blue-600 text-white"
                      : "bg-zinc-200 text-zinc-500"
                  }`}>
                    {!step1Done ? "🔒" : royaltyFile ? "✓" : "📑"}
                  </div>
                  <div className="text-left overflow-hidden">
                    {royaltyFile ? (
                      <>
                        <p className="font-semibold text-zinc-900 text-xs truncate max-w-[260px] sm:max-w-[360px]">
                          {royaltyFile.name}
                        </p>
                        <p className="text-[10px] text-zinc-400">
                          {formatFileSize(royaltyFile.size)} • Click to change
                        </p>
                      </>
                    ) : (
                      <>
                        <p className="font-semibold text-zinc-800 text-xs">
                          {step1Done ? "Select Franchise Royalty Calculator File" : "Step 1 required first"}
                        </p>
                        <p className="text-[10px] text-zinc-400">
                          {step1Done ? "Drop Franchise Royalty Calculator.xlsx" : "Locked until Step 1 completes"}
                        </p>
                      </>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Step 2 Action Button */}
          <div className="mt-2.5">
            <button
              onClick={handleProcessStep2}
              disabled={!step1Done || !royaltyFile || isProcessingStep2}
              className={`w-full py-2 px-3 rounded-lg font-semibold text-xs sm:text-sm flex items-center justify-center gap-2 transition-all shadow-xs ${
                !step1Done || !royaltyFile || isProcessingStep2
                  ? "bg-zinc-200 text-zinc-400 cursor-not-allowed"
                  : step2Done
                  ? "bg-blue-700 hover:bg-blue-800 text-white cursor-pointer"
                  : "bg-blue-600 hover:bg-blue-700 text-white cursor-pointer active:scale-[0.99]"
              }`}
            >
              {isProcessingStep2 ? (
                <>
                  <svg className="animate-spin h-3.5 w-3.5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"></path>
                  </svg>
                  Updating Calculations Sheet...
                </>
              ) : step2Done ? (
                `✓ Step 2 Complete (${step2FileName || "Re-download"})`
              ) : step1Stats?.latestDateStr ? (
                `⚡ Step 2: Process & Download Royalty Calculator ${step1Stats.latestDateStr}`
              ) : (
                "⚡ Step 2: Process & Download Royalty Calculator"
              )}
            </button>
          </div>
        </section>

        {/* Footer info banner */}
        <footer className="px-3 py-1.5 bg-zinc-200/50 rounded-lg text-[11px] text-zinc-500 flex items-center justify-between shrink-0">
          <span>Tableau Data 5 columns injected: Date, MOD#, Store, Gross Sales, Net Sales (Franchisee omitted)</span>
          <span className="font-mono">Sheet4 / Table11</span>
        </footer>

      </div>
    </div>
  );
}
