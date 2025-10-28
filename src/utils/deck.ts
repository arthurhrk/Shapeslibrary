import { spawn } from "child_process";
import { existsSync, mkdirSync } from "fs";
import { join } from "path";
import { tmpdir } from "os";
import { getLibraryRoot } from "./paths";

export function getDeckPath(): string {
  const root = getLibraryRoot();
  return join(root, "library_deck.pptx");
}

export async function ensureDeck(): Promise<string> {
  const deck = getDeckPath();
  const dir = join(deck, "..\n");
  if (!existsSync(dir)) {
    try { mkdirSync(dir, { recursive: true }); } catch {}
  }
  if (existsSync(deck)) return deck;
  // Create a blank deck via PowerPoint
  await runPs(`
$ErrorActionPreference = "Stop"
try {
  $created = $false
  try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
  $app.DisplayAlerts = 0
  $p = $app.Presentations.Add(0)
  $p.Slides.Add(1,12) | Out-Null
  $p.SaveAs('${deck.replace(/'/g, "''")}',24)
  $p.Close()
  'OK'
} catch { "ERROR:$($_.Exception.Message)"; exit 1 }
`);
  return deck;
}

export async function addShapeToDeckFromPptx(sourcePptx: string): Promise<number> {
  const deck = await ensureDeck();
  const out = await runPs(`
$ErrorActionPreference = "Stop"
try {
  $deck='${deck.replace(/'/g, "''")}'; $src='${sourcePptx.replace(/'/g, "''")}';
  $created = $false
  try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
  $app.DisplayAlerts = 0
  $d = $app.Presentations.Open($deck,$true,$false,$false)
  $s = $d.Slides.Add($d.Slides.Count+1,12)
  $p = $app.Presentations.Open($src,$true,$false,$false)
  $p.Slides.Item(1).Shapes.Range().Copy()
  $s.Shapes.Paste() | Out-Null
  $p.Close()
  $d.Save()
  $idx = $s.SlideIndex
  "OK:$idx"
} catch { "ERROR:$($_.Exception.Message)"; exit 1 }
`);
  const m = /^OK:(\d+)/.exec(out.trim());
  if (!m) throw new Error(`Failed to add to deck: ${out}`);
  return parseInt(m[1], 10);
}

export async function copyFromDeckToClipboard(slideIndex: number): Promise<void> {
  const deck = await ensureDeck();
  await runPs(`
$ErrorActionPreference = "Stop"
try {
  $deck='${deck.replace(/'/g, "''")}'; $idx=${slideIndex}
  $created = $false
  try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $app = New-Object -ComObject PowerPoint.Application; $created = $true }
  $app.DisplayAlerts = 0
  $d = $app.Presentations.Open($deck,$true,$false,$false)
  $sl = $d.Slides.Item($idx)
  if ($sl.Shapes.Count -gt 0) { $sl.Shapes.Range().Copy() }
  $d.Close()
  'OK'
} catch { "ERROR:$($_.Exception.Message)"; exit 1 }
`);
}

export async function insertFromDeckIntoActive(slideIndex: number): Promise<void> {
  const deck = await ensureDeck();
  await runPs(`
$ErrorActionPreference = "Stop"
try {
  $deck='${deck.replace(/'/g, "''")}'; $idx=${slideIndex}
  $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
  $app.DisplayAlerts = 0
  if ($app.Presentations.Count -eq 0) { "ERROR:No presentation is open"; exit 1 }
  $dest = $app.ActiveWindow.View.Slide
  if ($null -eq $dest) { $dest = $app.ActivePresentation.Slides.Item(1) }
  $d = $app.Presentations.Open($deck,$true,$false,$false)
  $sl = $d.Slides.Item($idx)
  if ($sl.Shapes.Count -gt 0) { $sl.Shapes.Range().Copy() }
  $dest.Shapes.Paste() | Out-Null
  $d.Close()
  'OK'
} catch { "ERROR:$($_.Exception.Message)"; exit 1 }
`);
}

async function runPs(script: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const temp = join(tmpdir(), `deck-${Date.now()}.ps1`);
    try { require("fs").writeFileSync(temp, script, "utf-8"); } catch (e) { return reject(e as Error); }
    const ps = spawn("powershell", ["-NoProfile","-NonInteractive","-ExecutionPolicy","Bypass","-File", temp]);
    let stdout = ""; let stderr = "";
    ps.stdout.on("data", (d) => stdout += d.toString());
    ps.stderr.on("data", (d) => stderr += d.toString());
    ps.on("error", (e) => done(e));
    ps.on("close", (code) => done(code === 0 ? null : new Error(`PowerShell failed (${code}). ${stderr || stdout}`)));
    function done(err: Error | null) {
      try { require("fs").unlinkSync(temp); } catch {}
      if (err) return reject(err);
      if (stdout.trim().startsWith("ERROR:")) return reject(new Error(stdout.trim().slice(6)));
      resolve(stdout);
    }
  });
}

