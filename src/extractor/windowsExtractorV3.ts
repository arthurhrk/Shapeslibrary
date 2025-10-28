/**
 * Windows PowerPoint shape extractor V3
 * Using spawn for real-time output and better debugging
 */

import { spawn } from "child_process";
import { writeFileSync, unlinkSync, existsSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { ExtractedShape, ExtractionResult } from "./types";

/**
 * Extract selected shape - V3 with spawn for real-time output
 */
export async function extractSelectedShapeWindows(): Promise<ExtractionResult> {
  // MUCH simpler script to test
  const script = `
try {
    Write-Host "STEP1: Getting PowerPoint"
    $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    Write-Host "STEP2: PowerPoint found"

    Write-Host "STEP3: Getting selection"
    $selection = $ppt.ActiveWindow.Selection
    Write-Host "STEP4: Selection type: $($selection.Type)"

    if ($selection.Type -ne 2) {
        Write-Output "ERROR:No shape selected (type: $($selection.Type))"
        exit 1
    }

    Write-Host "STEP5: Getting shape"
    $shape = $selection.ShapeRange.Item(1)
    Write-Host "STEP6: Shape name: $($shape.Name)"

    Write-Host "STEP7: Building data"
    $data = @{
        name = $shape.Name
        type = [int]$shape.AutoShapeType
        left = [math]::Round($shape.Left / 72, 3)
        top = [math]::Round($shape.Top / 72, 3)
        width = [math]::Round($shape.Width / 72, 3)
        height = [math]::Round($shape.Height / 72, 3)
        rotation = 0
    }

    Write-Host "STEP8: Converting to JSON"
    $json = $data | ConvertTo-Json -Compress
    Write-Host "STEP9: Outputting JSON"
    Write-Output $json
    Write-Host "STEP10: Done"

} catch {
    Write-Host "ERROR_CAUGHT: $($_.Exception.Message)"
    Write-Output "ERROR:$($_.Exception.Message)"
    exit 1
}
`;

  const tempScriptPath = join(tmpdir(), `raycast-capture-${Date.now()}.ps1`);

  return new Promise((resolve) => {
    console.log("=== V3 Extractor with Real-time Output ===");
    console.log("Creating script:", tempScriptPath);

    try {
      writeFileSync(tempScriptPath, script, "utf-8");
      console.log("Script written successfully");
    } catch (writeError) {
      console.error("Failed to write script:", writeError);
      resolve({
        success: false,
        error: `Failed to create temp script: ${writeError}`,
      });
      return;
    }

    console.log("Spawning PowerShell process...");

    const ps = spawn("powershell", [
      "-NoProfile",
      "-NonInteractive",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      tempScriptPath,
    ]);

    let stdout = "";
    let stderr = "";
    const stepLogs: string[] = [];

    ps.stdout.on("data", (data) => {
      const text = data.toString();
      stdout += text;
      console.log("[PowerShell STDOUT]:", text.trim());

      // Track steps
      if (text.includes("STEP")) {
        stepLogs.push(text.trim());
      }
    });

    ps.stderr.on("data", (data) => {
      const text = data.toString();
      stderr += text;
      console.error("[PowerShell STDERR]:", text.trim());
    });

    ps.on("error", (error) => {
      console.error("[PowerShell SPAWN ERROR]:", error);
      cleanup();
      resolve({
        success: false,
        error: `Failed to spawn PowerShell: ${error.message}`,
      });
    });

    ps.on("close", (code) => {
      console.log(`[PowerShell CLOSE] Exit code: ${code}`);
      console.log(`[PowerShell CLOSE] Steps completed: ${stepLogs.length}/10`);
      console.log(`[PowerShell CLOSE] Last step: ${stepLogs[stepLogs.length - 1] || "none"}`);

      cleanup();

      if (code !== 0) {
        console.error("PowerShell exited with error code:", code);
        resolve({
          success: false,
          error: `PowerShell failed with code ${code}. Stderr: ${stderr}`,
        });
        return;
      }

      // Parse output
      const output = stdout.trim();
      console.log("[FINAL OUTPUT]:", output);

      // Check for errors
      if (output.startsWith("ERROR:")) {
        resolve({
          success: false,
          error: output.replace("ERROR:", ""),
        });
        return;
      }

      // Find JSON in output (might have other lines)
      const lines = output.split("\n");
      const jsonLine = lines.find((line) => line.trim().startsWith("{"));

      if (!jsonLine) {
        console.error("No JSON found in output");
        console.error("Full output:", output);
        resolve({
          success: false,
          error: "No JSON data in PowerShell output",
        });
        return;
      }

      try {
        const data = JSON.parse(jsonLine);

        const shape: ExtractedShape = {
          name: data.name || "Unnamed",
          type: data.type || 1,
          position: { x: data.left || 1, y: data.top || 1 },
          size: { width: data.width || 2, height: data.height || 2 },
          rotation: data.rotation || 0,
          fill: {},
          line: { weight: 1 },
        };

        resolve({
          success: true,
          shape,
        });
      } catch (parseError) {
        console.error("JSON parse error:", parseError);
        console.error("Attempted to parse:", jsonLine);
        resolve({
          success: false,
          error: `Failed to parse JSON: ${parseError}`,
        });
      }
    });

    // Timeout after 10 seconds
    setTimeout(() => {
      console.warn("[TIMEOUT] PowerShell taking too long, killing process");
      console.warn(`[TIMEOUT] Steps completed: ${stepLogs.length}/10`);
      console.warn(`[TIMEOUT] Last step: ${stepLogs[stepLogs.length - 1] || "none"}`);

      ps.kill();
      cleanup();

      resolve({
        success: false,
        error: `Timeout after 10s. Last step: ${stepLogs[stepLogs.length - 1] || "none"}`,
      });
    }, 10000);

    function cleanup() {
      try {
        if (existsSync(tempScriptPath)) {
          unlinkSync(tempScriptPath);
          console.log("Temp script cleaned up");
        }
      } catch (err) {
        console.warn("Cleanup failed:", err);
      }
    }
  });
}
