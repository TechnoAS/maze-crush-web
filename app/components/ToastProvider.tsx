"use client";

import React, {
  createContext, useContext, useState, useCallback, useEffect, useRef,
} from "react";

// ── Types ──────────────────────────────────────────────────────────────────
export type ToastType = "success" | "error" | "warning" | "info" | "loading";

interface ToastItem {
  id: string;
  type: ToastType;
  title: string;
  message?: string;
  duration: number; // ms; 0 = persistent
  exiting?: boolean;
}

export interface ToastAPI {
  success: (title: string, message?: string, duration?: number) => string;
  error:   (title: string, message?: string, duration?: number) => string;
  warning: (title: string, message?: string, duration?: number) => string;
  info:    (title: string, message?: string, duration?: number) => string;
  loading: (title: string, message?: string) => string;
  dismiss: (id: string) => void;
  update:  (id: string, patch: Partial<Pick<ToastItem, "type" | "title" | "message" | "duration">>) => void;
}

// ── Context ────────────────────────────────────────────────────────────────
const ToastContext = createContext<ToastAPI | null>(null);

export function useToast(): ToastAPI {
  const ctx = useContext(ToastContext);
  if (!ctx) throw new Error("useToast must be used inside ToastProvider");
  return ctx;
}

// ── ID generator ───────────────────────────────────────────────────────────
let _n = 0;
const uid = () => `toast_${++_n}_${Date.now()}`;

// ── Provider ───────────────────────────────────────────────────────────────
export function ToastProvider({ children }: { children: React.ReactNode }) {
  const [toasts, setToasts] = useState<ToastItem[]>([]);

  const dismiss = useCallback((id: string) => {
    setToasts(p => p.map(t => t.id === id ? { ...t, exiting: true } : t));
    setTimeout(() => setToasts(p => p.filter(t => t.id !== id)), 380);
  }, []);

  const add = useCallback((type: ToastType, title: string, message?: string, duration = 4500): string => {
    const id = uid();
    setToasts(p => [{ id, type, title, message, duration }, ...p].slice(0, 5));
    if (duration > 0) setTimeout(() => dismiss(id), duration);
    return id;
  }, [dismiss]);

  const api: ToastAPI = {
    success: (t, m, d)  => add("success", t, m, d),
    error:   (t, m, d)  => add("error",   t, m, d ?? 6000),
    warning: (t, m, d)  => add("warning", t, m, d),
    info:    (t, m, d)  => add("info",    t, m, d),
    loading: (t, m)     => add("loading", t, m, 0),
    dismiss,
    update:  (id, patch) => setToasts(p => p.map(t => t.id === id ? { ...t, ...patch } : t)),
  };

  return (
    <ToastContext.Provider value={api}>
      {children}
      <div className="toast-stack" aria-live="polite" aria-label="Notifications">
        {toasts.map(t => (
          <ToastCard key={t.id} toast={t} onDismiss={() => dismiss(t.id)} />
        ))}
      </div>
    </ToastContext.Provider>
  );
}

// ── Config per type ────────────────────────────────────────────────────────
const CFG: Record<ToastType, { color: string; icon: React.ReactNode }> = {
  success: {
    color: "#22c55e",
    icon: (
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#22c55e" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
        <polyline points="22 4 12 14.01 9 11.01"/>
      </svg>
    ),
  },
  error: {
    color: "#ef4444",
    icon: (
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#ef4444" strokeWidth="2.5" strokeLinecap="round">
        <circle cx="12" cy="12" r="10"/>
        <line x1="15" y1="9" x2="9" y2="15"/>
        <line x1="9" y1="9" x2="15" y2="15"/>
      </svg>
    ),
  },
  warning: {
    color: "#f59e0b",
    icon: (
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#f59e0b" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
        <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
        <line x1="12" y1="9" x2="12" y2="13"/>
        <line x1="12" y1="17" x2="12.01" y2="17"/>
      </svg>
    ),
  },
  info: {
    color: "#3b8bff",
    icon: (
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#3b8bff" strokeWidth="2.5" strokeLinecap="round">
        <circle cx="12" cy="12" r="10"/>
        <line x1="12" y1="8" x2="12" y2="12"/>
        <line x1="12" y1="16" x2="12.01" y2="16"/>
      </svg>
    ),
  },
  loading: {
    color: "#9b6dff",
    icon: (
      <div style={{
        width: 20, height: 20,
        border: "2.5px solid rgba(155,109,255,0.25)",
        borderTop: "2.5px solid #9b6dff",
        borderRadius: "50%",
        animation: "spin 0.7s linear infinite",
        flexShrink: 0,
      }} />
    ),
  },
};

// ── Toast Card ─────────────────────────────────────────────────────────────
function ToastCard({ toast, onDismiss }: { toast: ToastItem; onDismiss: () => void }) {
  const [progress, setProgress] = useState(100);
  const rafRef = useRef<number>(0);
  const startRef = useRef(Date.now());
  const cfg = CFG[toast.type];

  useEffect(() => {
    if (toast.duration === 0) return;
    startRef.current = Date.now();
    const tick = () => {
      const pct = Math.max(0, 100 - ((Date.now() - startRef.current) / toast.duration) * 100);
      setProgress(pct);
      if (pct > 0) rafRef.current = requestAnimationFrame(tick);
    };
    rafRef.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(rafRef.current);
  }, [toast.duration]);

  return (
    <div
      className={`toast-card toast-${toast.type} ${toast.exiting ? "toast-exit" : "toast-enter"}`}
      style={{ "--tc": cfg.color } as React.CSSProperties}
      role="alert"
    >
      {/* Glow blob */}
      <div className="toast-glow" />

      <div className="toast-icon">{cfg.icon}</div>

      <div className="toast-body">
        <p className="toast-title">{toast.title}</p>
        {toast.message && <p className="toast-msg">{toast.message}</p>}
      </div>

      <button className="toast-close" onClick={onDismiss} aria-label="Dismiss notification">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round">
          <line x1="18" y1="6" x2="6" y2="18"/>
          <line x1="6" y1="6" x2="18" y2="18"/>
        </svg>
      </button>

      {toast.duration > 0 && (
        <div className="toast-progress-track">
          <div className="toast-progress-fill" style={{ width: `${progress}%` }} />
        </div>
      )}
    </div>
  );
}
