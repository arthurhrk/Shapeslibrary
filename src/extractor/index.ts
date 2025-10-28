/**
 * Cross-platform PowerPoint shape extractor
 */

import { platform } from "os";
import { extractSelectedShapeWindows } from "./windowsExtractor"; // Unified Windows extractor (spawn)
import { extractSelectedShapeMac } from "./macExtractor";
import { ExtractionResult } from "./types";

/**
 * Extract selected shape from PowerPoint (cross-platform)
 */
export async function captureShapeFromPowerPoint(): Promise<ExtractionResult> {
  const os = platform();

  console.log(`Detecting platform: ${os}`);

  if (os === "win32") {
    console.log("Using Windows extractor (V3 - spawn with real-time output)");
    return await extractSelectedShapeWindows();
  } else if (os === "darwin") {
    console.log("Using Mac extractor");
    return await extractSelectedShapeMac();
  } else {
    return {
      success: false,
      error: `Unsupported operating system: ${os}. Only Windows and macOS are supported.`,
    };
  }
}

// Export types and platform-specific extractors for advanced use
export * from "./types";
export { extractSelectedShapeWindows } from "./windowsExtractor";
export { extractSelectedShapeMac } from "./macExtractor";
