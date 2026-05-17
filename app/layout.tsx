import type { Metadata } from "next";
import "./globals.css";
import { ToastProvider } from "./components/ToastProvider";

export const metadata: Metadata = {
  title: "MazeCrush — WhatsApp Production Report Parser",
  description:
    "Instantly parse WhatsApp maize crushing daily production reports into structured data. Extract 80+ fields covering crushing, boiler, fuel, utilities, and vat metrics.",
  keywords: ["maize crushing", "production report", "WhatsApp parser", "Power BI", "Excel automation"],
  openGraph: {
    title: "MazeCrush — Production Report Parser",
    description: "Paste your WhatsApp production report and instantly extract all KPIs.",
    type: "website",
  },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>
        <ToastProvider>{children}</ToastProvider>
      </body>
    </html>
  );
}
