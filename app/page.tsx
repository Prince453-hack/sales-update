"use client";

import React, { useState, useRef, ChangeEvent, DragEvent } from "react";

interface UploadStats {
  updatedSheet: string;
  brinkRows: number;
  tableauRows: number;
  totalRows: number;
  reconciledStores: number;
  fileName: string;
  downloadUrl: string;
}

const masterSheets = [
  { id: 1, name: "Instructions" },
  { id: 2, name: "Summary" },
  { id: 3, name: "Brink Sales Summary by Location", badge: "Brink Target" },
  { id: 4, name: "Sheet2" },
  { id: 5, name: "Sheet5" },
  { id: 6, name: "Tableau Data", badge: "Tableau Target" },
  { id: 7, name: "GC Promo" },
  { id: 8, name: "Lookup" },
  { id: 9, name: "Rates from Tracker", badge: "Franchisee Source" },
];

export default function Home() {
  const [brinkFile, setBrinkFile] = useState<File | null>(null);
  const [tableauFile, setTableauFile] = useState<File | null>(null);

  const [isDraggingBrink, setIsDraggingBrink] = useState(false);
  const [isDraggingTableau, setIsDraggingTableau] = useState(false);

  const [isProcessing, setIsProcessing] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [stats, setStats] = useState<UploadStats | null>(null);

  const brinkInputRef = useRef<HTMLInputElement>(null);
  const tableauInputRef = useRef<HTMLInputElement>(null);

  const validateExcelFile = (file: File): boolean => {
    const lower = file.name.toLowerCase();
    return lower.endsWith(".xlsx") || lower.endsWith(".xls");
  };

  const handleBrinkChange = (e: ChangeEvent<HTMLInputElement>) => {
    setErrorMessage(null);
    setStats(null);
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (validateExcelFile(file)) {
        setBrinkFile(file);
      } else {
        setErrorMessage(
          "Please select a valid Excel file (.xlsx or .xls) for Brink.",
        );
      }
    }
  };

  const handleTableauChange = (e: ChangeEvent<HTMLInputElement>) => {
    setErrorMessage(null);
    setStats(null);
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (validateExcelFile(file)) {
        setTableauFile(file);
      } else {
        setErrorMessage(
          "Please select a valid Excel file (.xlsx or .xls) for Tableau.",
        );
      }
    }
  };

  const handleBrinkDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDraggingBrink(false);
    setErrorMessage(null);
    setStats(null);
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
    setStats(null);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (validateExcelFile(file)) setTableauFile(file);
      else setErrorMessage("Please select a valid Excel file for Tableau.");
    }
  };

  const handleProcessUpload = async () => {
    if (!brinkFile && !tableauFile) {
      setErrorMessage(
        "Please upload at least one file (Brink Sales Summary or Tableau Data).",
      );
      return;
    }

    setIsProcessing(true);
    setErrorMessage(null);
    setStats(null);

    try {
      const formData = new FormData();
      if (brinkFile) formData.append("brinkFile", brinkFile);
      if (tableauFile) formData.append("tableauFile", tableauFile);

      const response = await fetch("/api/update-sales", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let errText = "Failed to process files.";
        try {
          const json = await response.json();
          if (json.error) errText = json.error;
        } catch {
          errText = await response.text();
        }
        throw new Error(errText);
      }

      const updatedSheet =
        response.headers.get("X-Updated-Sheet") || "Franchise Sales Workbook";
      const brinkRows = parseInt(
        response.headers.get("X-Updated-Brink-Rows") || "0",
        10,
      );
      const tableauRows = parseInt(
        response.headers.get("X-Updated-Tableau-Rows") || "0",
        10,
      );
      const totalRows = parseInt(
        response.headers.get("X-Updated-Rows") || "0",
        10,
      );
      const reconciledStores = parseInt(
        response.headers.get("X-Reconciled-Stores") || "0",
        10,
      );

      let downloadFileName = "Franchise Sales (Updated).xlsx";
      const contentDisposition = response.headers.get("Content-Disposition");
      if (contentDisposition) {
        const match = contentDisposition.match(/filename="?([^";]+)"?/);
        if (match && match[1]) {
          downloadFileName = match[1];
        }
      }

      const blob = await response.blob();
      const downloadUrl = window.URL.createObjectURL(blob);

      const downloadLink = document.createElement("a");
      downloadLink.href = downloadUrl;
      downloadLink.download = downloadFileName;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);

      setStats({
        updatedSheet,
        brinkRows,
        tableauRows,
        totalRows,
        reconciledStores,
        fileName: downloadFileName,
        downloadUrl,
      });
    } catch (err) {
      setErrorMessage(
        err instanceof Error ? err.message : "Error processing file upload.",
      );
    } finally {
      setIsProcessing(false);
    }
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  const getButtonText = () => {
    if (isProcessing) return "Processing & Merging Sheets...";
    if (brinkFile && tableauFile)
      return "⚡ Update Both Sheets & Download Franchise Sales";
    if (brinkFile) return "⚡ Update Brink Sheet (#3) & Download";
    if (tableauFile) return "⚡ Update Tableau Sheet (#6) & Download";
    return "⚡ Select File(s) to Update & Download";
  };

  return (
    <div className="flex flex-col flex-1 items-center justify-start bg-zinc-100/70 p-6 md:p-12 font-sans text-zinc-900 min-h-screen">
      <main className="w-full max-w-3xl space-y-6">
        {/* Header Card */}
        <div className="bg-white rounded-2xl shadow-sm border border-zinc-200 p-6 sm:p-8">
          <div className="flex items-center gap-3 border-b border-zinc-100 pb-5 mb-6">
            <span className="flex h-11 w-11 items-center justify-center rounded-xl bg-emerald-50 text-emerald-600 font-semibold text-xl border border-emerald-100 shadow-xs">
              📊
            </span>
            <div>
              <h1 className="text-xl font-bold tracking-tight text-zinc-900">
                Franchise Sales Updater
              </h1>
              <p className="text-xs text-zinc-500 font-mono mt-0.5">
                Template: Franchise Sales 8.10.26 - 8.16.26.xlsx
              </p>
            </div>
          </div>

          <p className="text-sm text-zinc-600 leading-relaxed mb-6">
            Upload your{" "}
            <span className="font-semibold text-zinc-800">
              Brink Sales Summary
            </span>{" "}
            and/or{" "}
            <span className="font-semibold text-zinc-800">Tableau Data</span>{" "}
            Excel files.
          </p>

          {/* Dual Upload Section */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            {/* 1. Brink Upload Dropzone */}
            <div className="flex flex-col">
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold text-zinc-700 uppercase tracking-wider flex items-center gap-1.5">
                  <span>📊</span> 1. Brink Sales Summary
                </span>
                <span className="text-[11px] bg-zinc-100 text-zinc-600 px-2 py-0.5 rounded font-mono">
                  Sheet #3
                </span>
              </div>

              <input  
                ref={brinkInputRef}
                type="file"
                accept=".xlsx, .xls"
                onChange={handleBrinkChange}
                className="hidden"
                id="brink-file-input"
              />

              <div
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsDraggingBrink(true);
                }}
                onDragLeave={() => setIsDraggingBrink(false)}
                onDrop={handleBrinkDrop}
                onClick={() => brinkInputRef.current?.click()}
                className={`flex-1 border-2 border-dashed rounded-xl p-5 text-center cursor-pointer transition-all ${
                  isDraggingBrink
                    ? "border-emerald-500 bg-emerald-50/50 scale-[0.99]"
                    : brinkFile
                      ? "border-emerald-300 bg-emerald-50/30"
                      : "border-zinc-300 hover:border-zinc-400 bg-zinc-50/70 hover:bg-zinc-50"
                }`}
              >
                <div className="flex flex-col items-center justify-center gap-2">
                  <div
                    className={`w-10 h-10 rounded-full flex items-center justify-center text-lg ${
                      brinkFile
                        ? "bg-emerald-600 text-white"
                        : "bg-zinc-100 text-zinc-500"
                    }`}
                  >
                    {brinkFile ? "✓" : "📁"}
                  </div>
                  {brinkFile ? (
                    <div className="w-full">
                      <p className="font-semibold text-zinc-900 text-xs truncate px-2">
                        {brinkFile.name}
                      </p>
                      <p className="text-[11px] text-zinc-500 mt-0.5">
                        {formatFileSize(brinkFile.size)} • Click to change
                      </p>
                    </div>
                  ) : (
                    <div>
                      <p className="font-semibold text-zinc-800 text-xs">
                        Upload Brink File
                      </p>
                      <p className="text-[11px] text-zinc-400 mt-0.5">
                        Drop .xlsx / .xls
                      </p>
                    </div>
                  )}
                </div>
              </div>

              {brinkFile && (
                <button
                  type="button"
                  onClick={() => {
                    setBrinkFile(null);
                    if (brinkInputRef.current) brinkInputRef.current.value = "";
                  }}
                  className="mt-1.5 text-[11px] text-red-500 hover:text-red-700 self-end font-medium cursor-pointer"
                >
                  Remove Brink file
                </button>
              )}
            </div>

            {/* 2. Tableau Upload Dropzone */}
            <div className="flex flex-col">
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold text-zinc-700 uppercase tracking-wider flex items-center gap-1.5">
                  <span>📈</span> 2. Tableau Data
                </span>
                <span className="text-[11px] bg-emerald-50 text-emerald-700 font-semibold px-2 py-0.5 rounded font-mono">
                  Sheet #6
                </span>
              </div>

              <input
                ref={tableauInputRef}
                type="file"
                accept=".xlsx, .xls"
                onChange={handleTableauChange}
                className="hidden"
                id="tableau-file-input"
              />

              <div
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsDraggingTableau(true);
                }}
                onDragLeave={() => setIsDraggingTableau(false)}
                onDrop={handleTableauDrop}
                onClick={() => tableauInputRef.current?.click()}
                className={`flex-1 border-2 border-dashed rounded-xl p-5 text-center cursor-pointer transition-all ${
                  isDraggingTableau
                    ? "border-emerald-500 bg-emerald-50/50 scale-[0.99]"
                    : tableauFile
                      ? "border-emerald-300 bg-emerald-50/30"
                      : "border-zinc-300 hover:border-zinc-400 bg-zinc-50/70 hover:bg-zinc-50"
                }`}
              >
                <div className="flex flex-col items-center justify-center gap-2">
                  <div
                    className={`w-10 h-10 rounded-full flex items-center justify-center text-lg ${
                      tableauFile
                        ? "bg-emerald-600 text-white"
                        : "bg-zinc-100 text-zinc-500"
                    }`}
                  >
                    {tableauFile ? "✓" : "📈"}
                  </div>
                  {tableauFile ? (
                    <div className="w-full">
                      <p className="font-semibold text-zinc-900 text-xs truncate px-2">
                        {tableauFile.name}
                      </p>
                      <p className="text-[11px] text-zinc-500 mt-0.5">
                        {formatFileSize(tableauFile.size)} • Click to change
                      </p>
                    </div>
                  ) : (
                    <div>
                      <p className="font-semibold text-zinc-800 text-xs">
                        Upload Tableau File
                      </p>
                      <p className="text-[11px] text-zinc-400 mt-0.5">
                        Auto-fills Franchisees
                      </p>
                    </div>
                  )}
                </div>
              </div>

              {tableauFile && (
                <button
                  type="button"
                  onClick={() => {
                    setTableauFile(null);
                    if (tableauInputRef.current)
                      tableauInputRef.current.value = "";
                  }}
                  className="mt-1.5 text-[11px] text-red-500 hover:text-red-700 self-end font-medium cursor-pointer"
                >
                  Remove Tableau file
                </button>
              )}
            </div>
          </div>

          {/* Action button */}
          <div className="mt-6 flex gap-3">
            <button
              onClick={handleProcessUpload}
              disabled={(!brinkFile && !tableauFile) || isProcessing}
              className={`flex-1 flex items-center justify-center gap-2 rounded-xl py-3.5 px-4 font-semibold text-sm transition-all shadow-xs ${
                (!brinkFile && !tableauFile) || isProcessing
                  ? "bg-zinc-200 text-zinc-400 cursor-not-allowed"
                  : "bg-emerald-600 hover:bg-emerald-700 text-white active:scale-[0.99] cursor-pointer shadow-emerald-600/20"
              }`}
            >
              {isProcessing ? (
                <>
                  <svg
                    className="animate-spin h-4 w-4 text-white"
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                  >
                    <circle
                      className="opacity-25"
                      cx="12"
                      cy="12"
                      r="10"
                      stroke="currentColor"
                      strokeWidth="4"
                    ></circle>
                    <path
                      className="opacity-75"
                      fill="currentColor"
                      d="M4 12a8 8 0 018-8v8H4z"
                    ></path>
                  </svg>
                  Processing & Merging Sheets...
                </>
              ) : (
                getButtonText()
              )}
            </button>

            {(brinkFile || tableauFile) && !isProcessing && (
              <button
                onClick={() => {
                  setBrinkFile(null);
                  setTableauFile(null);
                  setStats(null);
                  setErrorMessage(null);
                  if (brinkInputRef.current) brinkInputRef.current.value = "";
                  if (tableauInputRef.current)
                    tableauInputRef.current.value = "";
                }}
                className="px-4 py-3.5 rounded-xl border border-zinc-200 text-zinc-600 hover:bg-zinc-100 text-sm font-medium transition-colors cursor-pointer"
              >
                Clear All
              </button>
            )}
          </div>

          {/* Error Message */}
          {errorMessage && (
            <div className="mt-5 rounded-xl border border-red-200 bg-red-50 p-4 text-sm text-red-600">
              <div className="flex items-start gap-2">
                <span className="font-bold">⚠️</span>
                <div>
                  <p className="font-semibold">Unable to process file</p>
                  <p className="text-xs text-red-700 mt-1">{errorMessage}</p>
                </div>
              </div>
            </div>
          )}

          {/* Success / Stats Section */}
          {stats && (
            <div className="mt-6 rounded-xl border border-emerald-200 bg-emerald-50/60 p-5 space-y-4">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2 text-emerald-800 font-semibold text-sm">
                  <span>✅</span> Successfully merged and downloaded!
                </div>
                <a
                  href={stats.downloadUrl}
                  download={stats.fileName}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-emerald-700 hover:text-emerald-900 bg-emerald-100/80 hover:bg-emerald-100 px-3 py-1.5 rounded-lg transition-colors"
                >
                  <span>📥</span> Download Again
                </a>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 pt-1">
                <div className="bg-white rounded-lg p-3 border border-emerald-100 shadow-2xs">
                  <p className="text-xs text-zinc-500">Brink Stores Updated</p>
                  <p className="text-lg font-bold text-zinc-900 mt-0.5">
                    {stats.brinkRows}
                  </p>
                  <p className="text-[11px] text-zinc-400">
                    Sheet #3 locations
                  </p>
                </div>

                <div className="bg-white rounded-lg p-3 border border-emerald-100 shadow-2xs">
                  <p className="text-xs text-zinc-500">Tableau Rows Updated</p>
                  <p className="text-lg font-bold text-emerald-600 mt-0.5">
                    {stats.tableauRows}
                  </p>
                  <p className="text-[11px] text-zinc-400">
                    Sheet #6 (Franchisees filled)
                  </p>
                </div>

                <div className="bg-white rounded-lg p-3 border border-emerald-100 shadow-2xs col-span-2 sm:col-span-1">
                  <p className="text-xs text-zinc-500">Total Records Merged</p>
                  <p className="text-lg font-bold text-zinc-900 mt-0.5">
                    {stats.totalRows}
                  </p>
                  <p className="text-[11px] text-emerald-600 font-medium">
                    100% styles preserved
                  </p>
                </div>

                {stats.reconciledStores > 0 && (
                  <div className="bg-emerald-100/60 rounded-lg p-3 border border-emerald-200 shadow-2xs col-span-2 sm:col-span-3 flex items-center justify-between">
                    <div>
                      <p className="text-xs font-semibold text-emerald-950">
                        Variance (Net Sales) Auto-Reconciled
                      </p>
                      <p className="text-[11px] text-emerald-800 mt-0.5">
                        Deducted net sales variances &gt; $1 from Tableau Data for {stats.reconciledStores} store{stats.reconciledStores > 1 ? "s" : ""}, making Net Sales variances $0.
                      </p>
                    </div>
                    <span className="text-xs font-bold text-emerald-800 bg-emerald-200/80 px-2.5 py-1 rounded-full whitespace-nowrap">
                      {stats.reconciledStores} store{stats.reconciledStores > 1 ? "s" : ""} reconciled
                    </span>
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}
