"use client";
import dynamic from "next/dynamic";
import { useEffect, useState } from "react";
import { installStorageShim } from "../lib/storage-shim";

const MKIS = dynamic(() => import("../components/MKIS.jsx"), { ssr: false });

export default function Page() {
  const [ready, setReady] = useState(false);
  useEffect(() => { installStorageShim(); setReady(true); }, []);
  if (!ready) return null;
  return <MKIS />;
}
