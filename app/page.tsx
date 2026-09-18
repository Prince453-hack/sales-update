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
  // Current active step (1: Upload files, 2: Review sales, 3: Calculate royalty)
  const [currentStep, setCurrentStep] = useState<1 | 2 | 3>(1);

  // Step 1: Franchise Sales state
  const [brinkFile, setBrinkFile] = useState<File | null>(null);
  const [tableauFile, setTableauFile] = useState<File | null>(null);
  const [isDraggingBrink, setIsDraggingBrink] = useState(false);
  const [isDraggingTableau, setIsDraggingTableau] = useState(false);
  const [isProcessingStep1, setIsProcessingStep1] = useState(false);
  const [step1Done, setStep1Done] = useState(false);
  const [step1Stats, setStep1Stats] = useState<Step1Stats | null>(null);
  const [step1DownloadUrl, setStep1DownloadUrl] = useState<string | null>(null);
  const [step1FileName, setStep1FileName] = useState<string>(
    "Franchise Sales (Updated).xlsx"
  );
  const [reconciledTableauRows, setReconciledTableauRows] = useState<
    any[] | null
  >(null);

  // Step 2 & 3: Royalty Calculator state
  const [royaltyFile, setRoyaltyFile] = useState<File | null>(null);
  const [isDraggingRoyalty, setIsDraggingRoyalty] = useState(false);
  const [isProcessingStep2, setIsProcessingStep2] = useState(false);
  const [step2Done, setStep2Done] = useState(false);
  const [step2DownloadUrl, setStep2DownloadUrl] = useState<string | null>(null);
  const [step2FileName, setStep2FileName] = useState<string | null>(null);
  const [step2RecordsCount, setStep2RecordsCount] = useState<number | null>(
    null
  );

  // Status & Feedback
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  // Refs for hidden inputs
  const brinkInputRef = useRef<HTMLInputElement>(null);
  const tableauInputRef = useRef<HTMLInputElement>(null);
  const royaltyInputRef = useRef<HTMLInputElement>(null);

  const validateExcelFile = (file: File): boolean => {
    const lower = file.name.toLowerCase();
    return lower.endsWith(".xlsx") || lower.endsWith(".xls");
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  const handleBrinkChange = (e: ChangeEvent<HTMLInputElement>) => {
    setErrorMessage(null);
    setSuccessMessage(null);
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (validateExcelFile(file)) {
        setBrinkFile(file);
      } else {
        setErrorMessage("Please select a valid Excel file (.xlsx or .xls) for Brink.");
      }
    }
  };

  const handleTableauChange = (e: ChangeEvent<HTMLInputElement>) => {
    setErrorMessage(null);
    setSuccessMessage(null);
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (validateExcelFile(file)) {
        setTableauFile(file);
      } else {
        setErrorMessage("Please select a valid Excel file (.xlsx or .xls) for Tableau.");
      }
    }
  };

  const handleRoyaltyChange = (e: ChangeEvent<HTMLInputElement>) => {
    setErrorMessage(null);
    setSuccessMessage(null);
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (validateExcelFile(file)) {
        setRoyaltyFile(file);
      } else {
        setErrorMessage(
          "Please select a valid Excel file (.xlsx or .xls) for Royalty Calculator."
        );
      }
    }
  };

  const handleBrinkDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDraggingBrink(false);
    setErrorMessage(null);
    setSuccessMessage(null);
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
    setSuccessMessage(null);
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
    setSuccessMessage(null);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (validateExcelFile(file)) setRoyaltyFile(file);
      else setErrorMessage("Please select a valid Excel file for Royalty Calculator.");
    }
  };

  // Reconcile Brink + Tableau -> Franchise Sales (Updated).xlsx
  const handleProcessStep1 = async (autoDownload = false): Promise<any[] | null> => {
    if (!brinkFile || !tableauFile) {
      setErrorMessage(
        "Both Brink Sales Summary and Tableau Data files are required."
      );
      return null;
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
        let errText = "Failed to reconcile sales files.";
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

      // Prepare download URL
      const byteCharacters = atob(resData.file.data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: resData.file.mimeType });
      const downloadUrl = window.URL.createObjectURL(blob);

      if (autoDownload) {
        const downloadLink = document.createElement("a");
        downloadLink.href = downloadUrl;
        downloadLink.download = resData.file.fileName;
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
      }

      setStep1DownloadUrl(downloadUrl);
      setStep1FileName(resData.file.fileName);
      setStep1Stats(resData.stats);
      setReconciledTableauRows(resData.tableauRows);
      setStep1Done(true);
      return resData.tableauRows;
    } catch (err) {
      setErrorMessage(
        err instanceof Error ? err.message : "Error reconciling sales."
      );
      return null;
    } finally {
      setIsProcessingStep1(false);
    }
  };

  // Step 1: Proceed from Upload files to Review sales
  const handleProceedToReview = async () => {
    if (!brinkFile || !tableauFile) {
      setErrorMessage("Please select both Brink and Tableau files before continuing.");
      return;
    }

    if (!step1Done) {
      const rows = await handleProcessStep1(false);
      if (rows) {
        setCurrentStep(2);
      }
    } else {
      setCurrentStep(2);
    }
  };

  // Download reconciled sales handler
  const handleDownloadReconciledSales = () => {
    if (step1Done && step1DownloadUrl) {
      const downloadLink = document.createElement("a");
      downloadLink.href = step1DownloadUrl;
      downloadLink.download = step1FileName;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);
      setSuccessMessage("Reconciled sales downloaded.");
    } else {
      handleProcessStep1(true);
    }
  };

  // Step 3: Calculate royalty & download
  const handleProcessStep2 = async () => {
    let rowsToUse = reconciledTableauRows;

    if (!rowsToUse || rowsToUse.length === 0) {
      if (brinkFile && tableauFile) {
        rowsToUse = await handleProcessStep1(false);
        if (!rowsToUse || rowsToUse.length === 0) {
          return;
        }
      } else {
        setErrorMessage("Please complete sales reconciliation first.");
        return;
      }
    }

    if (!royaltyFile) {
      setErrorMessage(
        "Please select a Franchise Royalty Calculator file to upload (click Replace)."
      );
      if (royaltyInputRef.current) {
        royaltyInputRef.current.click();
      }
      return;
    }

    setIsProcessingStep2(true);
    setErrorMessage(null);
    setSuccessMessage(null);

    try {
      const formData = new FormData();
      formData.append("royaltyFile", royaltyFile);
      formData.append("tableauRows", JSON.stringify(rowsToUse));

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
        throw new Error(
          resData.error || "Failed to process Royalty Calculator."
        );
      }

      // Download final royalty workbook
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
      setStep2RecordsCount(
        resData.stats?.recordsCount || rowsToUse.length
      );
      setStep2Done(true);
      setSuccessMessage("Royalty calculator calculated and downloaded successfully!");
    } catch (err) {
      setErrorMessage(
        err instanceof Error ? err.message : "Error processing Royalty Calculator."
      );
    } finally {
      setIsProcessingStep2(false);
    }
  };

  const handleResetAll = () => {
    setCurrentStep(1);
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
    setSuccessMessage(null);
    if (brinkInputRef.current) brinkInputRef.current.value = "";
    if (tableauInputRef.current) tableauInputRef.current.value = "";
    if (royaltyInputRef.current) royaltyInputRef.current.value = "";
  };

  return (
    <div className="min-h-screen bg-[#f4f5f8] text-slate-800 flex flex-col items-center py-6 px-4 sm:px-6 antialiased selection:bg-[#9e0b2f]/15">
      {/* Hidden File Inputs */}
      <input
        ref={brinkInputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleBrinkChange}
        className="hidden"
      />
      <input
        ref={tableauInputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleTableauChange}
        className="hidden"
      />
      <input
        ref={royaltyInputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleRoyaltyChange}
        className="hidden"
      />

      <div className="w-full max-w-[860px] flex flex-col gap-6">
        {/* TOP BAR: Logo + Title (No date picker, No step pills) and Reset Button */}
        <header className="flex items-center justify-between w-full pt-1 pb-1">
          <div className="flex items-center gap-3">
            {/* Logo: 3 rounded vertical bars in crimson */}
            <div className="flex items-end gap-[3px] h-6 shrink-0">
              <div className="w-1.5 h-3 bg-[#9e0b2f] rounded-xs" />
              <div className="w-1.5 h-4.5 bg-[#9e0b2f] rounded-xs" />
              <div className="w-1.5 h-6 bg-[#9e0b2f] rounded-xs" />
            </div>
            <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-[#111827]">
              Franchise Sales &amp; Royalty Hub
            </h1>
          </div>

          {/* Reset Button */}
          <button
            onClick={handleResetAll}
            className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 hover:text-slate-900 transition-colors px-3 py-1.5 rounded-lg hover:bg-slate-200/60 cursor-pointer"
            title="Reset all files and start over"
          >
            <svg
              className="w-4 h-4 text-slate-500"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M21 2v6h-6" />
              <path d="M3 12a9 9 0 0 1 15-6.7L21 8" />
              <path d="M3 22v-6h6" />
              <path d="M21 12a9 9 0 0 1-15 6.7L3 16" />
            </svg>
            Reset
          </button>
        </header>

        {/* 3-STEP PROGRESS STEPPER */}
        <div className="flex items-center justify-center w-full py-1">
          <div className="flex items-center">
            {/* Step 1: 1. Upload files */}
            <button
              onClick={() => setCurrentStep(1)}
              className="flex items-center gap-2 cursor-pointer group"
            >
              <div
                className={`w-7 h-7 rounded-full flex items-center justify-center font-bold text-xs text-white shadow-xs transition-colors ${
                  step1Done || (brinkFile && tableauFile) || currentStep > 1
                    ? "bg-[#15803d]"
                    : currentStep === 1
                      ? "bg-[#9e0b2f]"
                      : "bg-slate-300"
                }`}
              >
                {step1Done || currentStep > 1 ? (
                  <svg
                    className="w-4 h-4"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="3"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <polyline points="20 6 9 17 4 12" />
                  </svg>
                ) : (
                  "1"
                )}
              </div>
              <span
                className={`text-sm font-bold whitespace-nowrap ${
                  currentStep === 1 ? "text-slate-900" : "text-slate-700"
                }`}
              >
                1. Upload files
              </span>
            </button>

            {/* Connecting line 1 -> 2 */}
            <div
              className={`h-[2.5px] w-10 sm:w-20 md:w-24 mx-2 sm:mx-3 transition-colors ${
                currentStep >= 2 || step1Done ? "bg-[#15803d]" : "bg-slate-200"
              }`}
            />

            {/* Step 2: 2. Review sales */}
            <button
              onClick={() => {
                if (brinkFile && tableauFile) setCurrentStep(2);
              }}
              disabled={!brinkFile || !tableauFile}
              className={`flex items-center gap-2 transition-opacity ${
                !brinkFile || !tableauFile ? "opacity-60 cursor-not-allowed" : "cursor-pointer"
              }`}
            >
              <div
                className={`w-7 h-7 rounded-full flex items-center justify-center font-bold text-xs text-white shadow-xs transition-colors ${
                  currentStep > 2 || (step1Done && currentStep === 3)
                    ? "bg-[#15803d]"
                    : currentStep === 2
                      ? "bg-[#9e0b2f]"
                      : "bg-slate-300"
                }`}
              >
                {currentStep > 2 ? (
                  <svg
                    className="w-4 h-4"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="3"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <polyline points="20 6 9 17 4 12" />
                  </svg>
                ) : (
                  "2"
                )}
              </div>
              <span
                className={`text-sm font-bold whitespace-nowrap ${
                  currentStep === 2 ? "text-slate-900" : "text-slate-700"
                }`}
              >
                2. Review sales
              </span>
            </button>

            {/* Connecting line 2 -> 3 */}
            <div
              className={`h-[2.5px] w-10 sm:w-20 md:w-24 mx-2 sm:mx-3 transition-colors ${
                currentStep === 3 ? "bg-[#9e0b2f]" : "bg-slate-200"
              }`}
            />

            {/* Step 3: 3. Calculate royalty */}
            <button
              onClick={() => {
                if (brinkFile && tableauFile) setCurrentStep(3);
              }}
              disabled={!brinkFile || !tableauFile}
              className={`flex items-center gap-2 transition-opacity ${
                !brinkFile || !tableauFile ? "opacity-60 cursor-not-allowed" : "cursor-pointer"
              }`}
            >
              <div
                className={`w-7 h-7 rounded-full flex items-center justify-center font-bold text-xs text-white shadow-xs transition-colors ${
                  step2Done
                    ? "bg-[#15803d]"
                    : currentStep === 3
                      ? "bg-[#9e0b2f]"
                      : "bg-slate-300"
                }`}
              >
                {step2Done ? (
                  <svg
                    className="w-4 h-4"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="3"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <polyline points="20 6 9 17 4 12" />
                  </svg>
                ) : (
                  "3"
                )}
              </div>
              <span
                className={`text-sm font-bold whitespace-nowrap ${
                  currentStep === 3 ? "text-slate-900" : "text-slate-700"
                }`}
              >
                3. Calculate royalty
              </span>
            </button>
          </div>
        </div>

        {/* FEEDBACK BANNERS */}
        {errorMessage && (
          <div className="w-full bg-[#fdf2f4] border border-[#fecdd3] text-[#9e0b2f] px-4 py-3 rounded-xl flex items-center justify-between text-sm shadow-xs animate-in fade-in duration-200">
            <div className="flex items-center gap-2">
              <svg
                className="w-4 h-4 shrink-0 text-[#9e0b2f]"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <circle cx="12" cy="12" r="10" />
                <line x1="12" y1="8" x2="12" y2="12" />
                <line x1="12" y1="16" x2="12.01" y2="16" />
              </svg>
              <span className="font-medium">{errorMessage}</span>
            </div>
            <button
              onClick={() => setErrorMessage(null)}
              className="text-xs font-bold text-[#9e0b2f]/70 hover:text-[#9e0b2f] cursor-pointer"
            >
              ✕
            </button>
          </div>
        )}

        {successMessage && (
          <div className="w-full bg-[#f0fdf4] border border-[#bbf7d0] text-[#15803d] px-4 py-3 rounded-xl flex items-center justify-between text-sm shadow-xs animate-in fade-in duration-200">
            <div className="flex items-center gap-2">
              <svg
                className="w-4 h-4 shrink-0 text-[#15803d]"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
                <polyline points="22 4 12 14.01 9 11.01" />
              </svg>
              <span className="font-medium">{successMessage}</span>
            </div>
            <button
              onClick={() => setSuccessMessage(null)}
              className="text-xs font-bold text-[#15803d]/70 hover:text-[#15803d] cursor-pointer"
            >
              ✕
            </button>
          </div>
        )}

        {/* MAIN CARD */}
        <div className="bg-white rounded-2xl border border-slate-200/90 shadow-sm p-6 sm:p-8 flex flex-col gap-6">
          {/* ========================================================================= */}
          {/* STEP 1: UPLOAD FILES (INITIAL VIEW) */}
          {/* ========================================================================= */}
          {currentStep === 1 && (
            <>
              {/* Header */}
              <div className="flex items-start gap-4">
                <div className="w-11 h-11 rounded-xl bg-[#fdf2f4] flex items-center justify-center shrink-0 border border-[#fecdd3]/60 shadow-2xs">
                  <svg
                    className="w-6 h-6 text-[#9e0b2f]"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                    <polyline points="17 8 12 3 7 8" />
                    <line x1="12" y1="3" x2="12" y2="15" />
                  </svg>
                </div>
                <div>
                  <h2 className="text-2xl font-bold tracking-tight text-slate-900 leading-snug">
                    Upload sales files
                  </h2>
                  <p className="text-sm text-slate-500 mt-0.5">
                    Upload Brink Sales Summary and Tableau Data files to start reconciliation.
                  </p>
                </div>
              </div>

              {/* Upload boxes for Brink and Tableau */}
              <div className="flex flex-col gap-3">
                <h3 className="text-sm font-bold text-slate-900 tracking-tight">
                  Sales files
                </h3>

                <div className="border border-slate-200/90 rounded-xl overflow-hidden divide-y divide-slate-100 bg-white shadow-2xs">
                  {/* Brink File Row / Dropzone */}
                  <div
                    onDragOver={(e) => {
                      e.preventDefault();
                      setIsDraggingBrink(true);
                    }}
                    onDragLeave={() => setIsDraggingBrink(false)}
                    onDrop={handleBrinkDrop}
                    onClick={() => brinkInputRef.current?.click()}
                    className={`p-4 flex items-center justify-between cursor-pointer transition-colors ${
                      isDraggingBrink ? "bg-rose-50/70 ring-2 ring-[#9e0b2f]" : "hover:bg-slate-50/70"
                    }`}
                  >
                    <div className="flex items-center gap-3.5 min-w-0">
                      {/* Excel Icon */}
                      <div className="relative flex items-center justify-center shrink-0 w-8 h-9">
                        <svg
                          viewBox="0 0 36 42"
                          fill="none"
                          xmlns="http://www.w3.org/2000/svg"
                          className="w-full h-full drop-shadow-2xs"
                        >
                          <path
                            d="M4 2C2.89543 2 2 2.89543 2 4V38C2 39.1046 2.89543 40 4 40H32C33.1046 40 34 39.1046 34 38V12L24 2H4Z"
                            fill="#F8FAFC"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                          <path
                            d="M24 2V12H34"
                            fill="#E2E8F0"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                        </svg>
                        <div className="absolute -left-1 bottom-0.5 w-4.5 h-4.5 rounded-sm bg-[#107c41] flex items-center justify-center shadow-xs">
                          <span className="text-white text-[10px] font-black leading-none">
                            X
                          </span>
                        </div>
                      </div>

                      <div className="min-w-0">
                        <h4 className="font-bold text-sm text-slate-900 leading-tight">
                          Brinks Sales Summary
                        </h4>
                        <p className="text-xs text-slate-500 truncate mt-0.5 font-normal">
                          {brinkFile
                            ? `${brinkFile.name} (${formatFileSize(brinkFile.size)})`
                            : "Click or drag & drop Brink Excel file (.xlsx)"}
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-3 shrink-0 ml-3">
                      {brinkFile ? (
                        <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-[#f0fdf4] text-[#15803d] border border-[#bbf7d0]">
                          <svg
                            className="w-3.5 h-3.5 text-[#15803d]"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="3"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <polyline points="20 6 9 17 4 12" />
                          </svg>
                          Ready
                        </span>
                      ) : (
                        <span className="text-xs font-semibold text-[#9e0b2f] bg-rose-50 px-2.5 py-1 rounded-md">
                          Select file
                        </span>
                      )}

                      {brinkFile && (
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            brinkInputRef.current?.click();
                          }}
                          className="flex items-center gap-1 text-xs font-semibold text-[#9e0b2f] hover:text-[#7f0925] hover:bg-rose-50 px-2 py-1 rounded-md transition-colors cursor-pointer"
                        >
                          Replace
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Tableau File Row / Dropzone */}
                  <div
                    onDragOver={(e) => {
                      e.preventDefault();
                      setIsDraggingTableau(true);
                    }}
                    onDragLeave={() => setIsDraggingTableau(false)}
                    onDrop={handleTableauDrop}
                    onClick={() => tableauInputRef.current?.click()}
                    className={`p-4 flex items-center justify-between cursor-pointer transition-colors ${
                      isDraggingTableau ? "bg-rose-50/70 ring-2 ring-[#9e0b2f]" : "hover:bg-slate-50/70"
                    }`}
                  >
                    <div className="flex items-center gap-3.5 min-w-0">
                      {/* Excel Icon */}
                      <div className="relative flex items-center justify-center shrink-0 w-8 h-9">
                        <svg
                          viewBox="0 0 36 42"
                          fill="none"
                          xmlns="http://www.w3.org/2000/svg"
                          className="w-full h-full drop-shadow-2xs"
                        >
                          <path
                            d="M4 2C2.89543 2 2 2.89543 2 4V38C2 39.1046 2.89543 40 4 40H32C33.1046 40 34 39.1046 34 38V12L24 2H4Z"
                            fill="#F8FAFC"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                          <path
                            d="M24 2V12H34"
                            fill="#E2E8F0"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                        </svg>
                        <div className="absolute -left-1 bottom-0.5 w-4.5 h-4.5 rounded-sm bg-[#107c41] flex items-center justify-center shadow-xs">
                          <span className="text-white text-[10px] font-black leading-none">
                            X
                          </span>
                        </div>
                      </div>

                      <div className="min-w-0">
                        <h4 className="font-bold text-sm text-slate-900 leading-tight">
                          Tableau Data
                        </h4>
                        <p className="text-xs text-slate-500 truncate mt-0.5 font-normal">
                          {tableauFile
                            ? `${tableauFile.name} (${formatFileSize(tableauFile.size)})`
                            : "Click or drag & drop Tableau Excel file (.xlsx)"}
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-3 shrink-0 ml-3">
                      {tableauFile ? (
                        <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-[#f0fdf4] text-[#15803d] border border-[#bbf7d0]">
                          <svg
                            className="w-3.5 h-3.5 text-[#15803d]"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="3"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <polyline points="20 6 9 17 4 12" />
                          </svg>
                          Ready
                        </span>
                      ) : (
                        <span className="text-xs font-semibold text-[#9e0b2f] bg-rose-50 px-2.5 py-1 rounded-md">
                          Select file
                        </span>
                      )}

                      {tableauFile && (
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            tableauInputRef.current?.click();
                          }}
                          className="flex items-center gap-1 text-xs font-semibold text-[#9e0b2f] hover:text-[#7f0925] hover:bg-rose-50 px-2 py-1 rounded-md transition-colors cursor-pointer"
                        >
                          Replace
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              {/* Privacy Notice Banner */}
              <div className="bg-[#fff5f6] border border-[#fecdd3]/70 rounded-xl px-4 py-3 flex items-center gap-3 text-xs sm:text-sm text-slate-700 shadow-2xs">
                <svg
                  className="w-5 h-5 text-[#9e0b2f] shrink-0"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
                  <path d="m9 12 2 2 4-4" />
                </svg>
                <span>
                  <strong className="text-slate-900 font-semibold">Privacy notice:</strong>{" "}
                  Uploaded files are processed temporarily to generate your output. No file data is stored after processing.
                </span>
              </div>

              {/* Step 1 Action Button */}
              <div className="pt-2 flex items-center justify-between">
                <p className="text-xs text-slate-400">
                  Step 1 of 3: Uploading files
                </p>
                <button
                  onClick={handleProceedToReview}
                  disabled={!brinkFile || !tableauFile || isProcessingStep1}
                  className="bg-[#9e0b2f] hover:bg-[#850927] active:bg-[#6f0720] text-white font-semibold text-sm px-6 py-3 rounded-xl shadow-xs hover:shadow transition-all duration-150 flex items-center gap-2 cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {isProcessingStep1 ? (
                    <>
                      <svg
                        className="animate-spin w-4 h-4 text-white"
                        viewBox="0 0 24 24"
                        fill="none"
                      >
                        <circle
                          className="opacity-25"
                          cx="12"
                          cy="12"
                          r="10"
                          stroke="currentColor"
                          strokeWidth="4"
                        />
                        <path
                          className="opacity-75"
                          fill="currentColor"
                          d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                        />
                      </svg>
                      Reconciling sales...
                    </>
                  ) : (
                    <>
                      Continue to Review sales
                      <svg
                        className="w-4 h-4"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <polyline points="9 18 15 12 9 6" />
                      </svg>
                    </>
                  )}
                </button>
              </div>
            </>
          )}

          {/* ========================================================================= */}
          {/* STEP 2: REVIEW SALES */}
          {/* ========================================================================= */}
          {currentStep === 2 && (
            <>
              {/* Header */}
              <div className="flex items-start gap-4">
                <div className="w-11 h-11 rounded-xl bg-[#fdf2f4] flex items-center justify-center shrink-0 border border-[#fecdd3]/60 shadow-2xs">
                  <svg
                    className="w-6 h-6 text-[#9e0b2f]"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                    <polyline points="14 2 14 8 20 8" />
                    <line x1="16" y1="13" x2="8" y2="13" />
                    <line x1="16" y1="17" x2="8" y2="17" />
                    <polyline points="10 9 9 9 8 9" />
                  </svg>
                </div>
                <div>
                  <h2 className="text-2xl font-bold tracking-tight text-slate-900 leading-snug">
                    Review sales
                  </h2>
                  <p className="text-sm text-slate-500 mt-0.5">
                    Review validated files and reconciled stores before calculating royalty.
                  </p>
                </div>
              </div>

              {/* Sales files container */}
              <div className="flex flex-col gap-3">
                <h3 className="text-sm font-bold text-slate-900 tracking-tight">
                  Sales files
                </h3>

                <div className="border border-slate-200/90 rounded-xl overflow-hidden divide-y divide-slate-100 bg-white shadow-2xs">
                  {/* Row 1: Brink */}
                  <div className="p-4 flex items-center justify-between">
                    <div className="flex items-center gap-3.5 min-w-0">
                      <div className="relative flex items-center justify-center shrink-0 w-8 h-9">
                        <svg
                          viewBox="0 0 36 42"
                          fill="none"
                          xmlns="http://www.w3.org/2000/svg"
                          className="w-full h-full drop-shadow-2xs"
                        >
                          <path
                            d="M4 2C2.89543 2 2 2.89543 2 4V38C2 39.1046 2.89543 40 4 40H32C33.1046 40 34 39.1046 34 38V12L24 2H4Z"
                            fill="#F8FAFC"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                          <path
                            d="M24 2V12H34"
                            fill="#E2E8F0"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                        </svg>
                        <div className="absolute -left-1 bottom-0.5 w-4.5 h-4.5 rounded-sm bg-[#107c41] flex items-center justify-center shadow-xs">
                          <span className="text-white text-[10px] font-black leading-none">
                            X
                          </span>
                        </div>
                      </div>

                      <div className="min-w-0">
                        <h4 className="font-bold text-sm text-slate-900 leading-tight">
                          Brinks Sales Summary
                        </h4>
                        <p className="text-xs text-slate-500 truncate mt-0.5 font-normal">
                          {brinkFile ? brinkFile.name : "Brinks 8.10.26 - 8.16.26.xlsx"}
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-4 shrink-0 ml-3">
                      <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-[#f0fdf4] text-[#15803d] border border-[#bbf7d0]">
                        <svg
                          className="w-3.5 h-3.5 text-[#15803d]"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="3"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <polyline points="20 6 9 17 4 12" />
                        </svg>
                        Validated
                      </span>

                      <button
                        onClick={() => brinkInputRef.current?.click()}
                        className="flex items-center gap-1 text-xs font-semibold text-[#9e0b2f] hover:text-[#7f0925] hover:bg-rose-50 px-2 py-1 rounded-md transition-colors cursor-pointer"
                      >
                        <svg
                          className="w-3.5 h-3.5"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2.5"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M21 2v6h-6" />
                          <path d="M3 12a9 9 0 0 1 15-6.7L21 8" />
                          <path d="M3 22v-6h6" />
                          <path d="M21 12a9 9 0 0 1-15 6.7L3 16" />
                        </svg>
                        Replace
                      </button>
                    </div>
                  </div>

                  {/* Row 2: Tableau */}
                  <div className="p-4 flex items-center justify-between">
                    <div className="flex items-center gap-3.5 min-w-0">
                      <div className="relative flex items-center justify-center shrink-0 w-8 h-9">
                        <svg
                          viewBox="0 0 36 42"
                          fill="none"
                          xmlns="http://www.w3.org/2000/svg"
                          className="w-full h-full drop-shadow-2xs"
                        >
                          <path
                            d="M4 2C2.89543 2 2 2.89543 2 4V38C2 39.1046 2.89543 40 4 40H32C33.1046 40 34 39.1046 34 38V12L24 2H4Z"
                            fill="#F8FAFC"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                          <path
                            d="M24 2V12H34"
                            fill="#E2E8F0"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                        </svg>
                        <div className="absolute -left-1 bottom-0.5 w-4.5 h-4.5 rounded-sm bg-[#107c41] flex items-center justify-center shadow-xs">
                          <span className="text-white text-[10px] font-black leading-none">
                            X
                          </span>
                        </div>
                      </div>

                      <div className="min-w-0">
                        <h4 className="font-bold text-sm text-slate-900 leading-tight">
                          Tableau Data
                        </h4>
                        <p className="text-xs text-slate-500 truncate mt-0.5 font-normal">
                          {tableauFile ? tableauFile.name : "Tableau 8.10.26 - 8.16.26.xlsx"}
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-4 shrink-0 ml-3">
                      <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-[#f0fdf4] text-[#15803d] border border-[#bbf7d0]">
                        <svg
                          className="w-3.5 h-3.5 text-[#15803d]"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="3"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <polyline points="20 6 9 17 4 12" />
                        </svg>
                        Validated
                      </span>

                      <button
                        onClick={() => tableauInputRef.current?.click()}
                        className="flex items-center gap-1 text-xs font-semibold text-[#9e0b2f] hover:text-[#7f0925] hover:bg-rose-50 px-2 py-1 rounded-md transition-colors cursor-pointer"
                      >
                        <svg
                          className="w-3.5 h-3.5"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2.5"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M21 2v6h-6" />
                          <path d="M3 12a9 9 0 0 1 15-6.7L21 8" />
                          <path d="M3 22v-6h6" />
                          <path d="M21 12a9 9 0 0 1-15 6.7L3 16" />
                        </svg>
                        Replace
                      </button>
                    </div>
                  </div>
                </div>

                {/* Bottom of 2 files: Stats + Download summary / reconciled sales button */}
                <div className="flex flex-wrap items-center justify-between gap-3 pt-1 px-1">
                  <div className="flex items-center gap-2 text-sm text-slate-600">
                    <svg
                      className="w-5 h-5 text-[#9e0b2f] shrink-0"
                      viewBox="0 0 24 24"
                      fill="currentColor"
                    >
                      <path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z" />
                    </svg>
                    <span className="font-bold text-slate-900">
                      {step1Stats?.reconciledStores ?? 20} stores matched
                    </span>
                    <span className="text-slate-300 mx-1">|</span>
                    <span className="font-medium text-slate-600">
                      0 exceptions
                    </span>
                  </div>

                  <button
                    onClick={handleDownloadReconciledSales}
                    disabled={isProcessingStep1}
                    className="flex items-center gap-1.5 text-xs sm:text-sm font-semibold text-[#9e0b2f] hover:text-[#7f0925] hover:bg-rose-50 px-3.5 py-1.5 rounded-lg border border-[#fecdd3] transition-colors cursor-pointer disabled:opacity-50"
                  >
                    {isProcessingStep1 ? (
                      "Processing..."
                    ) : (
                      <>
                        <svg
                          className="w-4 h-4 text-[#9e0b2f]"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2.2"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                          <polyline points="7 10 12 15 17 10" />
                          <line x1="12" y1="15" x2="12" y2="3" />
                        </svg>
                        Download reconciled sales
                      </>
                    )}
                  </button>
                </div>
              </div>

              {/* Privacy Notice */}
              <div className="bg-[#fff5f6] border border-[#fecdd3]/70 rounded-xl px-4 py-3 flex items-center gap-3 text-xs sm:text-sm text-slate-700 shadow-2xs">
                <svg
                  className="w-5 h-5 text-[#9e0b2f] shrink-0"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
                  <path d="m9 12 2 2 4-4" />
                </svg>
                <span>
                  <strong className="text-slate-900 font-semibold">Privacy notice:</strong>{" "}
                  Uploaded files are processed temporarily to generate your output. No file data is stored after processing.
                </span>
              </div>

              {/* Actions: Back & Proceed to Step 3 */}
              <div className="pt-2 flex items-center justify-between">
                <button
                  onClick={() => setCurrentStep(1)}
                  className="flex items-center gap-1.5 text-xs sm:text-sm font-semibold text-slate-600 hover:text-slate-900 px-3 py-2 rounded-lg hover:bg-slate-100 transition-colors cursor-pointer"
                >
                  <svg
                    className="w-4 h-4"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2.5"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <polyline points="15 18 9 12 15 6" />
                  </svg>
                  Back to upload
                </button>

                <button
                  onClick={() => setCurrentStep(3)}
                  className="bg-[#9e0b2f] hover:bg-[#850927] active:bg-[#6f0720] text-white font-semibold text-sm px-6 py-3 rounded-xl shadow-xs hover:shadow transition-all duration-150 flex items-center gap-2 cursor-pointer"
                >
                  Continue to Royalty calculation
                  <svg
                    className="w-4 h-4"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2.5"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <polyline points="9 18 15 12 9 6" />
                  </svg>
                </button>
              </div>
            </>
          )}

          {/* ========================================================================= */}
          {/* STEP 3: CALCULATE ROYALTY (THE EXACT SCREEN IN THE MOCKUP) */}
          {/* ========================================================================= */}
          {currentStep === 3 && (
            <>
              {/* Card Header: Calculator Icon + Title + Subtitle */}
              <div className="flex items-start gap-4">
                <div className="w-11 h-11 rounded-xl bg-[#fdf2f4] flex items-center justify-center shrink-0 border border-[#fecdd3]/60 shadow-2xs">
                  <svg
                    className="w-6 h-6 text-[#9e0b2f]"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <rect x="4" y="2" width="16" height="20" rx="2" />
                    <line x1="8" y1="6" x2="16" y2="6" />
                    <line x1="16" y1="14" x2="16" y2="18" />
                    <path d="M16 10h.01" />
                    <path d="M12 10h.01" />
                    <path d="M8 10h.01" />
                    <path d="M12 14h.01" />
                    <path d="M8 14h.01" />
                    <path d="M12 18h.01" />
                    <path d="M8 18h.01" />
                  </svg>
                </div>
                <div>
                  <h2 className="text-2xl font-bold tracking-tight text-slate-900 leading-snug">
                    Calculate royalty
                  </h2>
                  <p className="text-sm text-slate-500 mt-0.5">
                    Review validated files, then generate the royalty calculator.
                  </p>
                </div>
              </div>

              {/* SECTION 1: Sales files */}
              <div className="flex flex-col gap-3">
                <h3 className="text-sm font-bold text-slate-900 tracking-tight">
                  Sales files
                </h3>

                {/* Container for the two sales files */}
                <div className="border border-slate-200/90 rounded-xl overflow-hidden divide-y divide-slate-100 bg-white shadow-2xs">
                  {/* Row 1: Brinks Sales Summary */}
                  <div
                    onDragOver={(e) => {
                      e.preventDefault();
                      setIsDraggingBrink(true);
                    }}
                    onDragLeave={() => setIsDraggingBrink(false)}
                    onDrop={handleBrinkDrop}
                    className={`p-4 flex items-center justify-between transition-colors ${
                      isDraggingBrink ? "bg-rose-50/60 ring-2 ring-[#9e0b2f] inset-0" : "hover:bg-slate-50/50"
                    }`}
                  >
                    <div className="flex items-center gap-3.5 min-w-0">
                      {/* Excel Icon */}
                      <div className="relative flex items-center justify-center shrink-0 w-8 h-9">
                        <svg
                          viewBox="0 0 36 42"
                          fill="none"
                          xmlns="http://www.w3.org/2000/svg"
                          className="w-full h-full drop-shadow-2xs"
                        >
                          <path
                            d="M4 2C2.89543 2 2 2.89543 2 4V38C2 39.1046 2.89543 40 4 40H32C33.1046 40 34 39.1046 34 38V12L24 2H4Z"
                            fill="#F8FAFC"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                          <path
                            d="M24 2V12H34"
                            fill="#E2E8F0"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                        </svg>
                        <div className="absolute -left-1 bottom-0.5 w-4.5 h-4.5 rounded-sm bg-[#107c41] flex items-center justify-center shadow-xs">
                          <span className="text-white text-[10px] font-black leading-none">
                            X
                          </span>
                        </div>
                      </div>

                      <div className="min-w-0">
                        <h4 className="font-bold text-sm text-slate-900 leading-tight">
                          Brinks Sales Summary
                        </h4>
                        <p className="text-xs text-slate-500 truncate mt-0.5 font-normal">
                          {brinkFile
                            ? `${brinkFile.name} (${formatFileSize(brinkFile.size)})`
                            : "Brinks 8.10.26 - 8.16.26.xlsx"}
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-4 shrink-0 ml-3">
                      {/* Validated Badge */}
                      <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-[#f0fdf4] text-[#15803d] border border-[#bbf7d0]">
                        <svg
                          className="w-3.5 h-3.5 text-[#15803d]"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="3"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <polyline points="20 6 9 17 4 12" />
                        </svg>
                        Validated
                      </span>

                      {/* Replace Button */}
                      <button
                        onClick={() => brinkInputRef.current?.click()}
                        className="flex items-center gap-1 text-xs font-semibold text-[#9e0b2f] hover:text-[#7f0925] hover:bg-rose-50 px-2 py-1 rounded-md transition-colors cursor-pointer"
                      >
                        <svg
                          className="w-3.5 h-3.5"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2.5"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M21 2v6h-6" />
                          <path d="M3 12a9 9 0 0 1 15-6.7L21 8" />
                          <path d="M3 22v-6h6" />
                          <path d="M21 12a9 9 0 0 1-15 6.7L3 16" />
                        </svg>
                        Replace
                      </button>
                    </div>
                  </div>

                  {/* Row 2: Tableau Data */}
                  <div
                    onDragOver={(e) => {
                      e.preventDefault();
                      setIsDraggingTableau(true);
                    }}
                    onDragLeave={() => setIsDraggingTableau(false)}
                    onDrop={handleTableauDrop}
                    className={`p-4 flex items-center justify-between transition-colors ${
                      isDraggingTableau ? "bg-rose-50/60 ring-2 ring-[#9e0b2f] inset-0" : "hover:bg-slate-50/50"
                    }`}
                  >
                    <div className="flex items-center gap-3.5 min-w-0">
                      {/* Excel Icon */}
                      <div className="relative flex items-center justify-center shrink-0 w-8 h-9">
                        <svg
                          viewBox="0 0 36 42"
                          fill="none"
                          xmlns="http://www.w3.org/2000/svg"
                          className="w-full h-full drop-shadow-2xs"
                        >
                          <path
                            d="M4 2C2.89543 2 2 2.89543 2 4V38C2 39.1046 2.89543 40 4 40H32C33.1046 40 34 39.1046 34 38V12L24 2H4Z"
                            fill="#F8FAFC"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                          <path
                            d="M24 2V12H34"
                            fill="#E2E8F0"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                        </svg>
                        <div className="absolute -left-1 bottom-0.5 w-4.5 h-4.5 rounded-sm bg-[#107c41] flex items-center justify-center shadow-xs">
                          <span className="text-white text-[10px] font-black leading-none">
                            X
                          </span>
                        </div>
                      </div>

                      <div className="min-w-0">
                        <h4 className="font-bold text-sm text-slate-900 leading-tight">
                          Tableau Data
                        </h4>
                        <p className="text-xs text-slate-500 truncate mt-0.5 font-normal">
                          {tableauFile
                            ? `${tableauFile.name} (${formatFileSize(tableauFile.size)})`
                            : "Tableau 8.10.26 - 8.16.26.xlsx"}
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-4 shrink-0 ml-3">
                      {/* Validated Badge */}
                      <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-[#f0fdf4] text-[#15803d] border border-[#bbf7d0]">
                        <svg
                          className="w-3.5 h-3.5 text-[#15803d]"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="3"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <polyline points="20 6 9 17 4 12" />
                        </svg>
                        Validated
                      </span>

                      {/* Replace Button */}
                      <button
                        onClick={() => tableauInputRef.current?.click()}
                        className="flex items-center gap-1 text-xs font-semibold text-[#9e0b2f] hover:text-[#7f0925] hover:bg-rose-50 px-2 py-1 rounded-md transition-colors cursor-pointer"
                      >
                        <svg
                          className="w-3.5 h-3.5"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2.5"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M21 2v6h-6" />
                          <path d="M3 12a9 9 0 0 1 15-6.7L21 8" />
                          <path d="M3 22v-6h6" />
                          <path d="M21 12a9 9 0 0 1-15 6.7L3 16" />
                        </svg>
                        Replace
                      </button>
                    </div>
                  </div>
                </div>

                {/* Bottom of 2 files: Stores matched stats + Download summary/reconciled sales button */}
                <div className="flex flex-wrap items-center justify-between gap-3 pt-1 px-1">
                  <div className="flex items-center gap-2 text-sm text-slate-600">
                    <svg
                      className="w-5 h-5 text-[#9e0b2f] shrink-0"
                      viewBox="0 0 24 24"
                      fill="currentColor"
                    >
                      <path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z" />
                    </svg>
                    <span className="font-bold text-slate-900">
                      {step1Stats?.reconciledStores ?? 20} stores matched
                    </span>
                    <span className="text-slate-300 mx-1">|</span>
                    <span className="font-medium text-slate-600">
                      0 exceptions
                    </span>
                  </div>

                  {/* Download Reconciled Sales / Summary File Button in bottom of 2 files */}
                  <button
                    onClick={handleDownloadReconciledSales}
                    disabled={isProcessingStep1}
                    className="flex items-center gap-1.5 text-xs sm:text-sm font-semibold text-[#9e0b2f] hover:text-[#7f0925] hover:bg-rose-50 px-3.5 py-1.5 rounded-lg border border-[#fecdd3] transition-colors cursor-pointer disabled:opacity-50"
                  >
                    {isProcessingStep1 ? (
                      <>
                        <svg
                          className="animate-spin w-4 h-4 text-[#9e0b2f]"
                          viewBox="0 0 24 24"
                          fill="none"
                        >
                          <circle
                            className="opacity-25"
                            cx="12"
                            cy="12"
                            r="10"
                            stroke="currentColor"
                            strokeWidth="4"
                          />
                          <path
                            className="opacity-75"
                            fill="currentColor"
                            d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                          />
                        </svg>
                        Reconciling...
                      </>
                    ) : (
                      <>
                        <svg
                          className="w-4 h-4 text-[#9e0b2f]"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2.2"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                          <polyline points="7 10 12 15 17 10" />
                          <line x1="12" y1="15" x2="12" y2="3" />
                        </svg>
                        Download reconciled sales
                      </>
                    )}
                  </button>
                </div>
              </div>

              {/* SECTION 2: Franchise Royalty Calculator */}
              <div className="flex flex-col gap-3 pt-2">
                <div className="flex items-center justify-between">
                  <h3 className="text-sm font-bold text-slate-900 tracking-tight">
                    Franchise Royalty Calculator
                  </h3>
                  <span className="text-xs text-slate-400 font-medium">
                    Calculations Sheet · 5 columns
                  </span>
                </div>

                {/* When royaltyFile is NOT yet uploaded: Show clear upload dropzone */}
                {!royaltyFile ? (
                  <div
                    onDragOver={(e) => {
                      e.preventDefault();
                      setIsDraggingRoyalty(true);
                    }}
                    onDragLeave={() => setIsDraggingRoyalty(false)}
                    onDrop={handleRoyaltyDrop}
                    onClick={() => royaltyInputRef.current?.click()}
                    className={`border-2 border-dashed rounded-xl p-5 sm:p-6 flex flex-col sm:flex-row items-center justify-between gap-4 cursor-pointer transition-all ${
                      isDraggingRoyalty
                        ? "border-[#9e0b2f] bg-rose-50/70 ring-2 ring-[#9e0b2f]/20"
                        : "border-slate-300 hover:border-[#9e0b2f] hover:bg-slate-50/80 bg-white"
                    }`}
                  >
                    <div className="flex items-center gap-3.5 min-w-0">
                      {/* Excel Document icon */}
                      <div className="relative flex items-center justify-center shrink-0 w-9 h-10">
                        <svg
                          viewBox="0 0 36 42"
                          fill="none"
                          xmlns="http://www.w3.org/2000/svg"
                          className="w-full h-full drop-shadow-2xs"
                        >
                          <path
                            d="M4 2C2.89543 2 2 2.89543 2 4V38C2 39.1046 2.89543 40 4 40H32C33.1046 40 34 39.1046 34 38V12L24 2H4Z"
                            fill="#F8FAFC"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                          <path
                            d="M24 2V12H34"
                            fill="#E2E8F0"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                        </svg>
                        <div className="absolute -left-1 bottom-0.5 w-4.5 h-4.5 rounded-sm bg-[#9e0b2f] flex items-center justify-center shadow-xs">
                          <svg
                            className="w-3 h-3 text-white"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="3"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          >
                            <line x1="12" y1="5" x2="12" y2="19" />
                            <line x1="5" y1="12" x2="19" y2="12" />
                          </svg>
                        </div>
                      </div>

                      <div className="min-w-0 text-left">
                        <h4 className="font-bold text-sm text-slate-900 leading-tight">
                          Upload Franchise Royalty Calculator
                        </h4>
                        <p className="text-xs text-slate-500 truncate mt-1 font-normal">
                          Click to browse or drag &amp; drop Franchise Royalty Calculator.xlsx
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <span className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-full text-xs font-semibold bg-[#fdf2f4] text-[#9e0b2f] border border-[#fecdd3]">
                        Upload file (.xlsx)
                      </span>
                    </div>
                  </div>
                ) : (
                  /* When royaltyFile IS uploaded: Show validated ready row */
                  <div
                    onDragOver={(e) => {
                      e.preventDefault();
                      setIsDraggingRoyalty(true);
                    }}
                    onDragLeave={() => setIsDraggingRoyalty(false)}
                    onDrop={handleRoyaltyDrop}
                    className={`border border-slate-200/90 rounded-xl p-4 flex items-center justify-between bg-white shadow-2xs transition-colors ${
                      isDraggingRoyalty ? "bg-rose-50/60 ring-2 ring-[#9e0b2f]" : "hover:bg-slate-50/50"
                    }`}
                  >
                    <div className="flex items-center gap-3.5 min-w-0">
                      {/* Excel Icon */}
                      <div className="relative flex items-center justify-center shrink-0 w-8 h-9">
                        <svg
                          viewBox="0 0 36 42"
                          fill="none"
                          xmlns="http://www.w3.org/2000/svg"
                          className="w-full h-full drop-shadow-2xs"
                        >
                          <path
                            d="M4 2C2.89543 2 2 2.89543 2 4V38C2 39.1046 2.89543 40 4 40H32C33.1046 40 34 39.1046 34 38V12L24 2H4Z"
                            fill="#F8FAFC"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                          <path
                            d="M24 2V12H34"
                            fill="#E2E8F0"
                            stroke="#CBD5E1"
                            strokeWidth="1.6"
                          />
                        </svg>
                        <div className="absolute -left-1 bottom-0.5 w-4.5 h-4.5 rounded-sm bg-[#107c41] flex items-center justify-center shadow-xs">
                          <span className="text-white text-[10px] font-black leading-none">
                            X
                          </span>
                        </div>
                      </div>

                      <div className="min-w-0">
                        <h4 className="font-bold text-sm text-slate-900 leading-tight">
                          Franchise Royalty Calculator
                        </h4>
                        <p className="text-xs text-slate-500 truncate mt-0.5 font-normal">
                          {royaltyFile.name} ({formatFileSize(royaltyFile.size)}) · Calculations Sheet · 5 columns
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-4 shrink-0 ml-3">
                      {/* Ready Badge */}
                      <span className="inline-flex items-center gap-1.5 px-3.5 py-1 rounded-full text-xs font-semibold bg-[#fef2f4] text-[#be123c] border border-[#fecdd3]">
                        <svg
                          className="w-3.5 h-3.5 text-[#be123c]"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="3"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <polyline points="20 6 9 17 4 12" />
                        </svg>
                        Ready
                      </span>

                      {/* Replace Button */}
                      <button
                        type="button"
                        onClick={() => royaltyInputRef.current?.click()}
                        className="flex items-center gap-1 text-xs font-semibold text-[#9e0b2f] hover:text-[#7f0925] hover:bg-rose-50 px-2 py-1 rounded-md transition-colors cursor-pointer"
                      >
                        <svg
                          className="w-3.5 h-3.5"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2.5"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M21 2v6h-6" />
                          <path d="M3 12a9 9 0 0 1 15-6.7L21 8" />
                          <path d="M3 22v-6h6" />
                          <path d="M21 12a9 9 0 0 1-15 6.7L3 16" />
                        </svg>
                        Replace
                      </button>
                    </div>
                  </div>
                )}
              </div>

              {/* Privacy Notice Banner */}
              <div className="bg-[#fff5f6] border border-[#fecdd3]/70 rounded-xl px-4 py-3 flex items-center gap-3 text-xs sm:text-sm text-slate-700 shadow-2xs">
                <svg
                  className="w-5 h-5 text-[#9e0b2f] shrink-0"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
                  <path d="m9 12 2 2 4-4" />
                </svg>
                <span>
                  <strong className="text-slate-900 font-semibold">Privacy notice:</strong>{" "}
                  Uploaded files are processed temporarily to generate your output. No file data is stored after processing.
                </span>
              </div>

              {/* BOTTOM SECTION: ONLY 1 DOWNLOAD BUTTON */}
              <div className="pt-2 flex items-center justify-between">
                <button
                  onClick={() => setCurrentStep(2)}
                  className="flex items-center gap-1.5 text-xs sm:text-sm font-semibold text-slate-600 hover:text-slate-900 px-3 py-2 rounded-lg hover:bg-slate-100 transition-colors cursor-pointer"
                >
                  <svg
                    className="w-4 h-4"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2.5"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <polyline points="15 18 9 12 15 6" />
                  </svg>
                  Back to Review sales
                </button>

                {/* Single Primary Download Button */}
                <button
                  onClick={handleProcessStep2}
                  disabled={!royaltyFile || isProcessingStep2 || isProcessingStep1}
                  className={`font-semibold text-sm px-6 py-3 rounded-xl shadow-xs transition-all duration-150 flex items-center gap-2 ${
                    !royaltyFile
                      ? "bg-slate-200 text-slate-400 cursor-not-allowed"
                      : "bg-[#9e0b2f] hover:bg-[#850927] active:bg-[#6f0720] text-white hover:shadow cursor-pointer"
                  }`}
                >
                  {isProcessingStep2 ? (
                    <>
                      <svg
                        className="animate-spin w-4 h-4 text-white"
                        viewBox="0 0 24 24"
                        fill="none"
                      >
                        <circle
                          className="opacity-25"
                          cx="12"
                          cy="12"
                          r="10"
                          stroke="currentColor"
                          strokeWidth="4"
                        />
                        <path
                          className="opacity-75"
                          fill="currentColor"
                          d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                        />
                      </svg>
                      Processing royalty...
                    </>
                  ) : (
                    <>
                      <svg
                        className="w-4 h-4 text-white fill-white"
                        viewBox="0 0 24 24"
                      >
                        <path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z" />
                      </svg>
                      Calculate &amp; download
                    </>
                  )}
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
