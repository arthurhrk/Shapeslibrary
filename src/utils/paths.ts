import { environment, getPreferenceValues } from "@raycast/api";
import { existsSync, mkdirSync } from "fs";
import { join, isAbsolute } from "path";
import { homedir } from "os";

type Prefs = { libraryPath?: string };

function expandUserPath(p: string): string {
  if (!p) return p;
  let out = p.trim();
  if ((out.startsWith('"') && out.endsWith('"')) || (out.startsWith("'") && out.endsWith("'"))) {
    out = out.slice(1, -1);
  }
  // Expand %VAR% on Windows
  out = out.replace(/%([A-Za-z0-9_]+)%/g, (_, v) => process.env[v] || "");
  // Expand ~ to home
  if (out.startsWith("~")) {
    out = join(homedir(), out.slice(1));
  }
  // If relative, anchor to home
  if (!isAbsolute(out)) {
    out = join(homedir(), out);
  }
  return out;
}

export function getLibraryRoot(): string {
  const prefs = getPreferenceValues<Prefs>();
  const configured =
    prefs.libraryPath && prefs.libraryPath.trim().length > 0 ? expandUserPath(prefs.libraryPath) : undefined;
  const base = configured || environment.supportPath;
  try {
    if (!existsSync(base)) mkdirSync(base, { recursive: true });
  } catch (e) {
    // Fallback to supportPath if creation failed
    if (!existsSync(environment.supportPath)) {
      try {
        mkdirSync(environment.supportPath, { recursive: true });
      } catch {}
    }
    return environment.supportPath;
  }
  return base;
}

export function getShapesDir(): string {
  const dir = join(getLibraryRoot(), "shapes");
  try {
    if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
  } catch {}
  return dir;
}

export function getAssetsDir(): string {
  const dir = join(getLibraryRoot(), "assets");
  try {
    if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
  } catch {}
  return dir;
}

export function getNativeDir(): string {
  const dir = join(getLibraryRoot(), "native");
  try {
    if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
  } catch {}
  return dir;
}
