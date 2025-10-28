import { spawn } from "child_process";
import { existsSync, mkdirSync } from "fs";
import { join } from "path";
import { tmpdir } from "os";
import { ShapeInfo } from "../types/shapes";
import { getLibraryRoot } from "./paths";
import { updateShapeInLibrary } from "./shapeSaver";

/**
 * Generate a PNG preview for a shape and update its JSON entry with the preview path.
 * Windows only (uses PowerPoint COM). On other platforms, this is a no-op.
 * @returns Absolute path to the generated PNG
 */
export async function generatePreview(shape: ShapeInfo): Promise<string | null> {
  if (process.platform !== "win32") return null;

  const libRoot = getLibraryRoot();
  const outDir = join(libRoot, "assets", shape.category);
  try {
    if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });
  } catch {}

  const outPng = join(outDir, `${shape.id}.png`);

  // Determine PPTX source
  let srcPptx: string | null = null;
  let tempToDelete: string | null = null;
  if (shape.nativePptx) {
    srcPptx = join(libRoot, shape.nativePptx);
  }
  if (!srcPptx || !existsSync(srcPptx)) {
    // Generate a temp PPTX using the generator
    const { generateShapePptx } = await import("../generator/pptxGenerator");
    srcPptx = await generateShapePptx(shape);
    tempToDelete = srcPptx;
  }

  await exportPptxToPngWindows(srcPptx, outPng);

  // Update JSON preview path (relative to assets)
  const rel = `${shape.category}/${shape.id}.png`;
  try {
    updateShapeInLibrary(shape.id, shape.category as any, { preview: rel });
  } catch {}

  if (tempToDelete) {
    try {
      require("fs").unlinkSync(tempToDelete);
    } catch {}
  }

  return outPng;
}

async function exportPptxToPngWindows(pptxPath: string, pngPath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const script = `
$ErrorActionPreference = "Stop"
try {
  $pptPath = '${pptxPath.replace(/'/g, "''")}'
  $pngPath = '${pngPath.replace(/'/g, "''")}'
  $pngDir = Split-Path -Parent $pngPath
  if (-not (Test-Path $pngDir)) { New-Item -ItemType Directory -Force -Path $pngDir | Out-Null }

  $created = $false
  try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
  $app.DisplayAlerts = 0
  $pres = $app.Presentations.Open($pptPath, $true, $false, $false)
  $slide = $pres.Slides.Item(1)
  $slide.Export($pngPath, 'PNG', 1600, 900)
  $pres.Close()
  if ($created) { $app.Visible = $true }
  Write-Output 'OK'
} catch {
  Write-Output "ERROR:$($_.Exception.Message)"; exit 1
}
`;

    const temp = join(tmpdir(), `preview-${Date.now()}.ps1`);
    try {
      require("fs").writeFileSync(temp, script, "utf-8");
    } catch (e) {
      return reject(e as Error);
    }

    const ps = spawn("powershell", [
      "-NoProfile",
      "-NonInteractive",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      temp,
    ]);
    let stdout = "";
    let stderr = "";
    ps.stdout.on("data", (d) => (stdout += d.toString()));
    ps.stderr.on("data", (d) => (stderr += d.toString()));
    ps.on("error", (e) => done(e));
    ps.on("close", (code) => done(code === 0 ? null : new Error(`PowerShell failed (${code}). ${stderr || stdout}`)));

    function done(err: Error | null) {
      try {
        require("fs").unlinkSync(temp);
      } catch {}
      if (err) return reject(err);
      if (stdout.trim().startsWith("ERROR:")) return reject(new Error(stdout.trim().slice(6)));
      resolve();
    }
  });
}

