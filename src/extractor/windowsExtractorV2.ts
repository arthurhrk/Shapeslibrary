/**
 * Windows PowerPoint shape extractor using COM automation
 * V2 - Using temporary file approach for reliability
 */

import { exec } from "child_process";
import { promisify } from "util";
import { writeFileSync, unlinkSync, existsSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { ExtractedShape, ExtractionResult } from "./types";

const execAsync = promisify(exec);

/**
 * Extract selected shape from PowerPoint using PowerShell COM automation
 * This version uses a temporary .ps1 file to avoid command-line escaping issues
 */
export async function extractSelectedShapeWindows(): Promise<ExtractionResult> {
  const script = `
$ErrorActionPreference = "Stop"
try {
    # Get running PowerPoint instance
    try {
        $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    } catch {
        Write-Output "ERROR:PowerPoint is not running. Please open PowerPoint first."
        exit 1
    }

    if (!$ppt) {
        Write-Output "ERROR:PowerPoint is not running"
        exit 1
    }

    # Check if presentation is open
    if ($ppt.Presentations.Count -eq 0) {
        Write-Output "ERROR:No presentation is open"
        exit 1
    }

    # Get active selection
    try {
        $selection = $ppt.ActiveWindow.Selection
    } catch {
        Write-Output "ERROR:Could not access PowerPoint window. Make sure PowerPoint window is active."
        exit 1
    }

    # Check if shape is selected (Type 2 = ppSelectionShapes)
    if ($selection.Type -ne 2) {
        $typeMsg = "Selection type: $($selection.Type). Expected: 2 (shape). "
        if ($selection.Type -eq 1) {
            $typeMsg += "You have text selected. Please select a shape instead."
        } elseif ($selection.Type -eq 3) {
            $typeMsg += "You have a slide selected. Please select a shape on the slide."
        } else {
            $typeMsg += "Please select a single shape in PowerPoint."
        }
        Write-Output "ERROR:$typeMsg"
        exit 1
    }

    # Get the first selected shape
    $shape = $selection.ShapeRange.Item(1)

    # Initialize data object
    $data = @{
        name = $shape.Name
        type = [int]$shape.AutoShapeType
        left = [math]::Round($shape.Left / 72, 3)
        top = [math]::Round($shape.Top / 72, 3)
        width = [math]::Round($shape.Width / 72, 3)
        height = [math]::Round($shape.Height / 72, 3)
        rotation = [math]::Round($shape.Rotation, 2)
    }

    # Extract fill properties
    try {
        if ($shape.Fill.Visible -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
            if ($shape.Fill.ForeColor) {
                $rgb = $shape.Fill.ForeColor.RGB
                $r = ($rgb -band 0xFF).ToString("X2")
                $g = (($rgb -shr 8) -band 0xFF).ToString("X2")
                $b = (($rgb -shr 16) -band 0xFF).ToString("X2")
                $data['fillColor'] = "$r$g$b"
            }

            if ($shape.Fill.Transparency) {
                $data['fillTransparency'] = [math]::Round($shape.Fill.Transparency, 2)
            }
        }
    } catch {
        # Fill properties not available, continue
    }

    # Extract line properties
    try {
        if ($shape.Line.Visible -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
            if ($shape.Line.ForeColor) {
                $rgb = $shape.Line.ForeColor.RGB
                $r = ($rgb -band 0xFF).ToString("X2")
                $g = (($rgb -shr 8) -band 0xFF).ToString("X2")
                $b = (($rgb -shr 16) -band 0xFF).ToString("X2")
                $data['lineColor'] = "$r$g$b"
            }

            $data['lineWeight'] = [math]::Round($shape.Line.Weight, 2)

            if ($shape.Line.Transparency) {
                $data['lineTransparency'] = [math]::Round($shape.Line.Transparency, 2)
            }
        }
    } catch {
        # Line properties not available, continue
    }

    # Output as JSON
    $data | ConvertTo-Json -Compress

} catch {
    Write-Output "ERROR:$($_.Exception.Message)"
    exit 1
}
`;

  // Create temporary PowerShell script file
  const tempScriptPath = join(tmpdir(), `raycast-shape-capture-${Date.now()}.ps1`);

  try {
    console.log("=== Starting PowerShell extraction (V2 - File Method) ===");
    console.log("Creating temporary script:", tempScriptPath);

    // Write script to temporary file
    writeFileSync(tempScriptPath, script, "utf-8");

    console.log("Executing PowerShell script file...");

    // Execute PowerShell script from file
    const { stdout, stderr } = await execAsync(
      `powershell -NoProfile -ExecutionPolicy Bypass -File "${tempScriptPath}"`,
      {
        encoding: "utf-8",
        timeout: 5000, // 5 second timeout (reduced from 10)
      }
    );

    console.log("PowerShell execution completed");

    // Clean up temporary file
    try {
      if (existsSync(tempScriptPath)) {
        unlinkSync(tempScriptPath);
        console.log("Temporary script cleaned up");
      }
    } catch (cleanupError) {
      console.warn("Failed to cleanup temp script:", cleanupError);
    }

    // Debug logging
    console.log("=== PowerShell Execution Debug ===");
    console.log("stdout length:", stdout.length);
    console.log("stderr length:", stderr.length);

    if (stderr && stderr.trim()) {
      console.error("PowerShell stderr:", stderr);
    }

    const output = stdout.trim();

    // Log raw output for debugging
    console.log("PowerShell raw output:", JSON.stringify(output));
    console.log("First 100 chars:", output.substring(0, 100));

    // Check for error messages
    if (output.startsWith("ERROR:")) {
      return {
        success: false,
        error: output.replace("ERROR:", ""),
      };
    }

    // Check if output is empty
    if (!output) {
      return {
        success: false,
        error: "PowerShell script returned no output. Check PowerPoint is running and a shape is selected.",
      };
    }

    // Parse JSON output
    let data;
    try {
      data = JSON.parse(output);
    } catch (parseError) {
      console.error("Failed to parse JSON:", output);
      console.error("Parse error:", parseError);
      return {
        success: false,
        error: `Failed to parse shape data. PowerShell output: ${output.substring(0, 200)}`,
      };
    }

    // Map to ExtractedShape format
    const shape: ExtractedShape = {
      name: data.name || "Unnamed Shape",
      type: data.type || 1,
      position: {
        x: data.left || 1,
        y: data.top || 1,
      },
      size: {
        width: data.width || 2,
        height: data.height || 2,
      },
      rotation: data.rotation || 0,
      fill: {
        color: data.fillColor,
        transparency: data.fillTransparency,
      },
      line: {
        color: data.lineColor,
        weight: data.lineWeight || 1,
        transparency: data.lineTransparency,
      },
    };

    return {
      success: true,
      shape,
    };
  } catch (error) {
    // Clean up temporary file on error
    try {
      if (existsSync(tempScriptPath)) {
        unlinkSync(tempScriptPath);
      }
    } catch (cleanupError) {
      // Ignore cleanup errors
    }

    // Handle execution errors
    if (error instanceof Error) {
      const anyErr = error as any;
      const errStdout: string | undefined = anyErr?.stdout;
      const errStderr: string | undefined = anyErr?.stderr;

      console.error("Execution error:", error);
      if (errStdout) console.error("PowerShell stdout (on error):", errStdout);
      if (errStderr) console.error("PowerShell stderr (on error):", errStderr);

      if (error.message.includes("timeout")) {
        return {
          success: false,
          error: "PowerShell command timed out. Is PowerPoint responding?",
        };
      }

      if (error.message.includes("cannot be loaded because running scripts is disabled")) {
        return {
          success: false,
          error:
            "PowerShell execution policy prevents running scripts. This should not happen with -ExecutionPolicy Bypass.",
        };
      }

      // If PowerShell printed a helpful error, surface it
      if (errStdout && errStdout.trim().startsWith("ERROR:")) {
        return {
          success: false,
          error: errStdout.trim().replace(/^ERROR:/, ""),
        };
      }

      return {
        success: false,
        error: `Failed to extract shape: ${error.message}`,
      };
    }

    return {
      success: false,
      error: "Unknown error occurred while extracting shape",
    };
  }
}
