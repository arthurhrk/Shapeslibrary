/**
 * Mac PowerPoint shape extractor using AppleScript
 */

import { exec } from "child_process";
import { promisify } from "util";
import { ExtractedShape, ExtractionResult } from "./types";

const execAsync = promisify(exec);

/**
 * Extract selected shape from PowerPoint using AppleScript
 */
export async function extractSelectedShapeMac(): Promise<ExtractionResult> {
  const script = `
    tell application "Microsoft PowerPoint"
      if (count of presentations) = 0 then
        return "ERROR:No presentation is open"
      end if

      tell active presentation
        if (count of (get selection shapes)) = 0 then
          return "ERROR:No shape selected. Please select a shape in PowerPoint."
        end if

        set selectedShape to first item of (get selection shapes)

        -- Extract properties
        set shapeName to name of selectedShape
        set shapeType to shape type of selectedShape as integer

        -- Get position and size (in points, will convert to inches)
        set leftPos to left position of selectedShape
        set topPos to top position of selectedShape
        set shapeWidth to width of selectedShape
        set shapeHeight to height of selectedShape

        -- Get rotation
        set rotationAngle to rotation of selectedShape

        -- Convert points to inches (72 points = 1 inch)
        set leftInches to (leftPos / 72) as text
        set topInches to (topPos / 72) as text
        set widthInches to (shapeWidth / 72) as text
        set heightInches to (shapeHeight / 72) as text

        -- Return as delimited string (JSON is complex in AppleScript)
        return shapeName & "|" & shapeType & "|" & leftInches & "|" & topInches & "|" & widthInches & "|" & heightInches & "|" & rotationAngle
      end tell
    end tell
  `;

  try {
    // Execute AppleScript
    const { stdout, stderr } = await execAsync(`osascript -e '${script.replace(/'/g, "'\\''")}'`, {
      encoding: "utf-8",
      timeout: 10000, // 10 second timeout
    });

    if (stderr && stderr.trim()) {
      console.error("AppleScript stderr:", stderr);
    }

    const output = stdout.trim();

    // Check for error messages
    if (output.startsWith("ERROR:")) {
      return {
        success: false,
        error: output.replace("ERROR:", ""),
      };
    }

    // Parse delimited output
    const parts = output.split("|");

    if (parts.length < 7) {
      return {
        success: false,
        error: "Invalid output from AppleScript",
      };
    }

    const [name, typeStr, leftStr, topStr, widthStr, heightStr, rotationStr] = parts;

    // Map to ExtractedShape format
    const shape: ExtractedShape = {
      name: name || "Unnamed Shape",
      type: parseInt(typeStr) || 1,
      position: {
        x: parseFloat(leftStr) || 1,
        y: parseFloat(topStr) || 1,
      },
      size: {
        width: parseFloat(widthStr) || 2,
        height: parseFloat(heightStr) || 2,
      },
      rotation: parseFloat(rotationStr) || 0,
      fill: {
        // AppleScript has limited access to fill properties
        // Will use default colors
      },
      line: {
        // AppleScript has limited access to line properties
        weight: 1, // Default line weight
      },
    };

    return {
      success: true,
      shape,
    };
  } catch (error) {
    // Handle execution errors
    if (error instanceof Error) {
      if (error.message.includes("timeout")) {
        return {
          success: false,
          error: "AppleScript command timed out. Is PowerPoint responding?",
        };
      }

      if (error.message.includes("not running") || error.message.includes("Application isn't running")) {
        return {
          success: false,
          error: "PowerPoint is not running",
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
