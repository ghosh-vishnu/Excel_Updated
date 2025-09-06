import React, { useCallback, useMemo, useRef, useState } from "react";
import { motion, AnimatePresence } from "framer-motion";

type FileItem = {
  id: string;
  file: File;
  name: string;
  size: number;
  status: "pending" | "queued" | "converting" | "success" | "done" | "error";
  errorMessage?: string;
};

type ConversionResult = {
  downloadUrl: string;
  openUrl?: string;
};

// Make backend endpoints configurable without touching code
// const UPLOAD_PATH = (import.meta as any).env?.VITE_UPLOAD_PATH || "/api/upload/";
// const CONVERT_PATH = (import.meta as any).env?.VITE_CONVERT_PATH || "/api/convert/";
// const PROGRESS_PATH = (import.meta as any).env?.VITE_PROGRESS_PATH || "/api/progress/";
// const RESULT_PATH = (import.meta as any).env?.VITE_RESULT_PATH || "/api/result/";
const API_BASE =
  (import.meta as any).env?.VITE_API_BASE || "http://127.0.0.1:8000";
const UPLOAD_PATH = (import.meta as any).env?.VITE_UPLOAD_PATH || "/api/upload/";
const CONVERT_PATH = (import.meta as any).env?.VITE_CONVERT_PATH || "/api/convert/";
const PROGRESS_PATH = (import.meta as any).env?.VITE_PROGRESS_PATH || "/api/progress/";
const RESULT_PATH = (import.meta as any).env?.VITE_RESULT_PATH || "/api/result/";


async function uploadFolderToBackend(files: File[]): Promise<{ jobId: string }>{
  const formData = new FormData();
  files.forEach((file) => formData.append("files", file, (file as any).webkitRelativePath || file.name));
  const res = await fetch(`${API_BASE}${UPLOAD_PATH}` , { method: "POST", body: formData });
  if (!res.ok) throw new Error("Upload failed");
  return res.json();
}

async function startBackendConversion(jobId: string): Promise<{ started: boolean }>{
  const url = `${API_BASE}${CONVERT_PATH}?jobId=${encodeURIComponent(jobId)}`;
  const res = await fetch(url, { method: "POST" });
  if (!res.ok) throw new Error("Failed to start conversion");
  return res.json();
}

async function pollConversionProgress(jobId: string): Promise<{ progress: number; done: boolean; error?: string }>{
  const url = `${API_BASE}${PROGRESS_PATH}?jobId=${encodeURIComponent(jobId)}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error("Progress check failed");
  return res.json();
}

async function fetchConversionResult(jobId: string): Promise<ConversionResult>{
  // Prefer CSV text if user wants to open in a new tab easily; keep xlsx fallback by toggling query
  const endpointUrl = `${API_BASE}${RESULT_PATH}?jobId=${encodeURIComponent(jobId)}`;
  const res = await fetch(endpointUrl);
  if (!res.ok) throw new Error("Result not ready");
  const blob = await res.blob();
  const resultUrl = URL.createObjectURL(blob);
  return { downloadUrl: resultUrl, openUrl: resultUrl };
}

async function fetchConversionCsv(jobId: string): Promise<string> {
  const endpointUrl = `${API_BASE}${RESULT_PATH}?jobId=${encodeURIComponent(jobId)}&format=csv`;
  const res = await fetch(endpointUrl);
  if (!res.ok) throw new Error("CSV not ready");
  return res.text();
}

function bytesToReadable(size: number): string {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  if (size < 1024 * 1024 * 1024) return `${(size / (1024 * 1024)).toFixed(1)} MB`;
  return `${(size / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

export default function WordToExcel(): React.ReactElement {
  const [files, setFiles] = useState<FileItem[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [jobId, setJobId] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [progress, setProgress] = useState(0);
  const [statusMessage, setStatusMessage] = useState<string>("");
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [result, setResult] = useState<ConversionResult | null>(null);

  const inputRef = useRef<HTMLInputElement | null>(null);
  const pollingRef = useRef<number | null>(null);

  const hasFiles = files.length > 0;

  const acceptedExtensions = React.useMemo(() => [
    ".doc",
    ".docx",
    ".rtf",
    ".odt",
  ], []);

  const onFileInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList) return;
    const newFiles: FileItem[] = [];
    for (let i = 0; i < fileList.length; i += 1) {
      const file = fileList.item(i);
      if (!file) continue;
      const isAccepted = acceptedExtensions.some((ext) => file.name.toLowerCase().endsWith(ext));
      if (!isAccepted) continue;
      newFiles.push({
        id: `${file.name}-${file.size}-${file.lastModified}-${i}`,
        file,
        name: (file as any).webkitRelativePath || file.name,
        size: file.size,
        status: "pending",
      });
    }
    setFiles((prev) => [...prev, ...newFiles]);
  }, [acceptedExtensions]);

  const onDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const dt = e.dataTransfer;
    const items = dt.items;
    const collected: File[] = [];
    if (items && items.length > 0) {
      for (let i = 0; i < items.length; i += 1) {
        const item = items[i];
        if (!item) continue;
        const entry = (item as any).webkitGetAsEntry?.();
        if (entry && entry.isDirectory) {
          continue;
        } else {
          const file = item.getAsFile();
          if (file) collected.push(file);
        }
      }
    } else {
      if (dt.files && dt.files.length > 0) {
        for (let i = 0; i < dt.files.length; i += 1) {
          const f = dt.files.item(i);
          if (f) collected.push(f);
        }
      }
    }
    const filtered = collected.filter((f) => acceptedExtensions.some((ext) => f.name.toLowerCase().endsWith(ext)));
    const mapped: FileItem[] = filtered.map((file, idx) => ({
      id: `${file.name}-${file.size}-${file.lastModified}-${idx}`,
      file,
      name: file.name,
      size: file.size,
      status: "pending",
    }));
    setFiles((prev) => [...prev, ...mapped]);
  }, [acceptedExtensions]);

  const removeFile = useCallback((id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
  }, []);

  const updateFileStatus = useCallback((id: string, status: FileItem["status"], errorMessage?: string) => {
    setFiles((prev) => prev.map((f) => 
      f.id === id ? { ...f, status, errorMessage } : f
    ));
  }, []);

  const resetAll = useCallback(() => {
    setFiles([]);
    setJobId(null);
    setIsUploading(false);
    setIsConverting(false);
    setProgress(0);
    setStatusMessage("");
    setErrorMessage("");
    setResult(null);
    if (pollingRef.current) {
      window.clearInterval(pollingRef.current);
      pollingRef.current = null;
    }
  }, []);

  const beginConversion = useCallback(async () => {
    try {
      setErrorMessage("");
      setStatusMessage("Uploading files...");
      setIsUploading(true);
      
      // Set all files to converting status
      files.forEach((file) => {
        updateFileStatus(file.id, "converting");
      });
      
      const uploadRes = await uploadFolderToBackend(files.map((f) => f.file));
      setIsUploading(false);
      setJobId(uploadRes.jobId);

      setStatusMessage("Starting conversion...");
      setIsConverting(true);
      await startBackendConversion(uploadRes.jobId);

      setStatusMessage("Converting...");
      setProgress(5);

      // Track individual file processing
      let processedFiles = 0;
      const totalFiles = files.length;
      let lastProgress = 0;
      
      pollingRef.current = window.setInterval(async () => {
        try {
          const p = await pollConversionProgress(uploadRes.jobId);
          if (p.error) throw new Error(p.error);
          
          // Calculate progress more smoothly based on actual progress
          const currentProgress = Math.min(100, Math.max(0, p.progress));
          setProgress(currentProgress);
          
          // Update file statuses based on progress
          // Backend now provides progress from 5% to 85% for file processing
          // Map this to individual file completion
          if (currentProgress > lastProgress) {
            // Calculate how many files should be completed based on progress
            // Progress 5-85% represents file processing (80% of total progress)
            const fileProcessingProgress = Math.max(0, currentProgress - 5); // Remove initial 5%
            const fileProgressRatio = Math.min(1, fileProcessingProgress / 80); // 80% for file processing
            const expectedCompletedFiles = Math.floor(fileProgressRatio * totalFiles);
            
            // Mark files as success if they should be completed
            for (let i = processedFiles; i < expectedCompletedFiles && i < totalFiles; i++) {
              updateFileStatus(files[i].id, "success");
            }
            processedFiles = Math.max(processedFiles, expectedCompletedFiles);
            lastProgress = currentProgress;
          }
          
          if (p.done) {
            // Mark all remaining files as success
            files.forEach((file) => {
              if (file.status === "converting") {
                updateFileStatus(file.id, "success");
              }
            });
            
            if (pollingRef.current) {
              window.clearInterval(pollingRef.current);
              pollingRef.current = null;
            }
            setStatusMessage("Finalizing...");
            const res = await fetchConversionResult(uploadRes.jobId);
            setResult(res);
            setIsConverting(false);
            setStatusMessage("Conversion complete");
          }
        } catch (err: any) {
          if (pollingRef.current) {
            window.clearInterval(pollingRef.current);
            pollingRef.current = null;
          }
          setIsConverting(false);
          setErrorMessage(err?.message || "An error occurred during conversion.");
          setStatusMessage("");
        }
      }, 300); // Even more responsive updates
    } catch (err: any) {
      setIsUploading(false);
      setIsConverting(false);
      setStatusMessage("");
      setErrorMessage(err?.message || "Failed to start conversion.");
    }
  }, [files, updateFileStatus]);

  return (
    <div className="min-h-screen bg-gray-50 text-gray-800 dark:bg-gray-950 dark:text-gray-100">
      {/* Header */}
      <header className="sticky top-0 z-10 bg-white/70 dark:bg-gray-900/60 backdrop-blur border-b border-gray-100 dark:border-gray-800 shadow-sm">
        <div className="mx-auto max-w-3xl px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="h-9 w-9 rounded-xl bg-indigo-600 text-white grid place-items-center font-bold">W</div>
            <div className="leading-tight">
              <p className="text-2xl font-bold text-gray-900 dark:text-gray-100">Word → Excel Converter</p>
            </div>
          </div>
          <button
            type="button"
            aria-label="Toggle theme"
            className="hidden md:inline-flex items-center gap-2 rounded-full border border-gray-200 dark:border-gray-700 px-3 py-1.5 text-sm text-gray-600 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-800 transition-colors"
          >
            <span className="h-2.5 w-2.5 rounded-full bg-gray-400" />
            Theme
          </button>
        </div>
      </header>

      {/* Main */}
      <main className="mx-auto max-w-3xl p-6">
        <div className="mb-6 md:mb-6">
          <h1 className="text-2xl md:text-3xl font-bold text-gray-900 dark:text-gray-100">Convert Word documents to Excel</h1>
          <p className="mt-2 text-gray-600 dark:text-gray-400">Upload a folder of Word files and we will convert them into a single Excel file. Simple, fast, and secure.</p>
        </div>

        

        {/* Upload Section */}
        <section>
          <motion.div
            layout
            className={`rounded-xl border-2 border-dashed ${isDragging ? "border-indigo-500 bg-indigo-50 dark:bg-indigo-900/20" : "border-gray-300 bg-white dark:bg-gray-900"} shadow-md p-8 md:p-10 transition-colors`}
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={(e) => {
              e.preventDefault();
              setIsDragging(false);
            }}
            onDrop={onDrop}
          >
            <div className="flex flex-col items-center text-center gap-4">
              <div className={`h-14 w-14 rounded-2xl grid place-items-center ${isDragging ? "bg-indigo-100 text-indigo-600 dark:bg-indigo-900/40" : "bg-gray-100 text-gray-500 dark:bg-gray-800 dark:text-gray-300"}`}>
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-7 w-7">
                  <path d="M12 3a5 5 0 0 0-5 5v2H6a4 4 0 0 0 0 8h12a4 4 0 0 0 0-8h-1V8a5 5 0 0 0-5-5Zm-1 9V8a1 1 0 1 1 2 0v4h2.5a.75.75 0 0 1 .53 1.28l-3.5 3.5a.75.75 0 0 1-1.06 0l-3.5-3.5A.75.75 0 0 1 8.5 12H11Z" />
                </svg>
              </div>

              <div>
                <p className="text-base md:text-lg font-medium text-gray-900 dark:text-gray-100">Drop folder here or click to browse</p>
                <p className="text-sm text-gray-500 dark:text-gray-400">Accepted: DOC, DOCX, RTF, ODT</p>
              </div>

              {/* Action row: browse + start + reset */}
              <div className="flex flex-wrap items-center justify-center gap-4">
                <button
                  type="button"
                  onClick={() => inputRef.current?.click()}
                  className="inline-flex items-center gap-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-medium px-6 py-3 shadow focus-visible:outline focus-visible:outline-2 focus-visible:outline-indigo-600"
                >
                  Browse Folder
                </button>
                <button
                  type="button"
                  disabled={!hasFiles || isUploading || isConverting}
                  onClick={beginConversion}
                  className="inline-flex justify-center items-center gap-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-medium px-6 py-3 shadow disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {isUploading ? "Uploading..." : isConverting ? "Converting..." : "Start Conversion"}
                </button>
                <button
                  type="button"
                  onClick={resetAll}
                  className="inline-flex items-center gap-2 rounded-xl border border-gray-300 dark:border-gray-700 text-gray-600 dark:text-gray-300 px-6 py-3 font-medium hover:bg-gray-100 dark:hover:bg-gray-800"
                >
                  Reset
                </button>
              </div>

              <input
                ref={inputRef}
                type="file"
                multiple
                // @ts-expect-error - non-standard but widely supported in Chromium-based browsers
                webkitdirectory="true"
                directory="true"
                className="hidden"
                onChange={onFileInputChange}
              />
            </div>
          </motion.div>
        </section>

        {/* Progress & Download (directly below upload box) */}
        <AnimatePresence>
          {(isUploading || isConverting || progress > 0 || statusMessage || result) && (
            <motion.section
              layout
              initial={{ opacity: 0, y: 12 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -12 }}
              transition={{ duration: 0.2 }}
              className="mt-8"
              aria-live="polite"
            >
              <div className="rounded-2xl bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 shadow-sm p-6">
                <div className="mb-2 text-sm text-gray-600 dark:text-gray-400">{progress}%</div>
                <div
                  className="w-full h-3 bg-gray-200 dark:bg-gray-700 rounded-full overflow-hidden"
                  role="progressbar"
                  aria-valuemin={0}
                  aria-valuemax={100}
                  aria-valuenow={progress}
                >
                  <motion.div
                    className="h-full bg-indigo-600"
                    initial={{ width: "0%" }}
                    animate={{ width: `${progress}%` }}
                    transition={{ ease: "easeInOut", duration: 0.5 }}
                    style={{ borderRadius: 9999 }}
                  />
                </div>
                <div className="mt-3 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                  <div className="text-sm text-gray-500 dark:text-gray-400">
                    {statusMessage || (progress > 0 ? `Progress: ${progress}%` : "Idle")}
                  </div>
                  {result && (
                    <div className="flex items-center gap-3">
                      <a
                        href={result.downloadUrl}
                        download
                        className="inline-flex items-center gap-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-medium px-5 py-2.5 shadow"
                      >
                        Download Excel File
                      </a>
                      {result.openUrl && (
                        <a
                          href={result.openUrl}
                          target="_blank"
                          rel="noreferrer"
                          className="inline-flex items-center gap-2 rounded-xl border border-gray-300 dark:border-gray-700 text-gray-700 dark:text-gray-300 px-5 py-2.5 font-medium hover:bg-gray-100 dark:hover:bg-gray-800"
                        >
                          Open in New Tab
                        </a>
                      )}
                    </div>
                  )}
                </div>
              </div>
            </motion.section>
          )}
        </AnimatePresence>

        {/* Uploaded Files List */}
        <AnimatePresence>
          {hasFiles && (
            <motion.section
              layout
              initial={{ opacity: 0, y: 12 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -12 }}
              transition={{ duration: 0.2 }}
              className="mt-8"
            >
              <div className="rounded-2xl bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 shadow-sm">
                <div className="p-5 md:p-6 border-b border-gray-100 dark:border-gray-700 flex items-center justify-between">
                  <h2 className="text-lg font-semibold text-gray-900 dark:text-gray-100">Files</h2>
                  <span className="text-sm text-gray-500 dark:text-gray-400">{files.length} selected</span>
                </div>
                <div className="max-h-96 overflow-y-auto">
                  <ul className="divide-y divide-gray-100 dark:divide-gray-700">
                    {files.map((item) => (
                      <li key={item.id} className="px-5 md:px-6 py-2 flex items-center gap-4">
                        <div className={`h-9 w-9 rounded-lg grid place-items-center relative ${
                          item.status === "success" 
                            ? "bg-green-50 text-green-600 dark:bg-green-900/30 dark:text-green-400"
                            : item.status === "converting"
                            ? "bg-blue-50 text-blue-600 dark:bg-blue-900/30 dark:text-blue-400"
                            : item.status === "error"
                            ? "bg-red-50 text-red-600 dark:bg-red-900/30 dark:text-red-400"
                            : "bg-indigo-50 text-indigo-600 dark:bg-indigo-900/30 dark:text-indigo-400"
                        }`}>
                          {item.status === "success" ? (
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-5 w-5">
                              <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
                            </svg>
                          ) : item.status === "converting" ? (
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-5 w-5 animate-spin">
                              <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
                            </svg>
                          ) : (
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-5 w-5">
                              <path d="M6 2a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h9.5a2 2 0 0 0 2-2V8.5L13.5 2H6Zm7 1.5L18.5 9H13a.5.5 0 0 1-.5-.5V3.5Z" />
                            </svg>
                          )}
                        </div>
                        <div className="flex-1 min-w-0">
                          <p className="truncate font-medium text-gray-900 dark:text-gray-100">{item.name}</p>
                          <p className="text-sm text-gray-500 dark:text-gray-400">{bytesToReadable(item.size)}</p>
                        </div>
                        <div className="hidden sm:block">
                          <span className={`text-xs rounded-full px-2 py-1 border ${
                            item.status === "success" 
                              ? "border-green-200 dark:border-green-700 text-green-600 dark:text-green-400 bg-green-50 dark:bg-green-900/30"
                              : item.status === "converting"
                              ? "border-blue-200 dark:border-blue-700 text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-900/30"
                              : item.status === "error"
                              ? "border-red-200 dark:border-red-700 text-red-600 dark:text-red-400 bg-red-50 dark:bg-red-900/30"
                              : "border-gray-200 dark:border-gray-700 text-gray-600 dark:text-gray-300"
                          }`}>
                            {item.status === "success" ? "✓ Success" : item.status}
                          </span>
                        </div>
                        <button
                          type="button"
                          onClick={() => removeFile(item.id)}
                          className="ml-2 inline-flex items-center justify-center h-8 w-8 rounded-lg text-red-500 hover:text-red-600 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors"
                          aria-label={`Remove ${item.name}`}
                        >
                          <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-5 w-5">
                            <path d="M9 3a1 1 0 0 0-1 1v1H5.5a1 1 0 1 0 0 2H6v12a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V7h.5a1 1 0 1 0 0-2H16V4a1 1 0 0 0-1-1H9Zm2 4a1 1 0 0 0-1 1v9a1 1 0 1 0 2 0V8a1 1 0 0 0-1-1Zm4 0a1 1 0 0 0-1 1v9a1 1 0 1 0 2 0V8a1 1 0 0 0-1-1Z" />
                          </svg>
                        </button>
                      </li>
                    ))}
                  </ul>
                </div>
              </div>
            </motion.section>
          )}
        </AnimatePresence>

        

        {/* Error */}
        <AnimatePresence>
          {errorMessage && (
            <motion.section
              layout
              initial={{ opacity: 0, y: 12 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -12 }}
              transition={{ duration: 0.2 }}
              className="mt-8"
              aria-live="assertive"
            >
              <div className="rounded-2xl border border-red-200 dark:border-red-900 bg-red-50 dark:bg-red-950 text-red-800 dark:text-red-300 p-5 flex items-start gap-3">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className="h-5 w-5 mt-0.5">
                  <path d="M12 2a10 10 0 1 0 10 10A10.011 10.011 0 0 0 12 2Zm1 15h-2v-2h2Zm0-4h-2V7h2Z" />
                </svg>
                <div>
                  <p className="font-semibold">Conversion failed</p>
                  <p className="text-sm">{errorMessage}</p>
                </div>
              </div>
            </motion.section>
          )}
        </AnimatePresence>

        {/* Result card removed in favor of inline download above */}
      </main>
    </div>
  );
}


