/**
 * Script to generate preview images for all shapes
 *
 * This script:
 * 1. Reads all shape definitions from src/shapes/*.json
 * 2. Generates a PowerPoint file for each shape
 * 3. Exports each slide as an image (requires PowerPoint or conversion library)
 * 4. Saves images to assets/ directory
 *
 * Usage: npm run generate-previews
 */

import pptxgen from "pptxgenjs";
import { readFileSync, writeFileSync, readdirSync, existsSync, mkdirSync } from "fs";
import { join, basename } from "path";

interface ShapeInfo {
  id: string;
  name: string;
  category: string;
  preview: string;
  pptxDefinition: any;
}

const SHAPES_DIR = join(__dirname, "../src/shapes");
const ASSETS_DIR = join(__dirname, "../assets");

/**
 * Load all shapes from JSON files
 */
function loadAllShapes(): ShapeInfo[] {
  const shapes: ShapeInfo[] = [];

  const jsonFiles = readdirSync(SHAPES_DIR).filter(f => f.endsWith(".json"));

  for (const file of jsonFiles) {
    const filePath = join(SHAPES_DIR, file);
    const content = readFileSync(filePath, "utf-8");
    const categoryShapes: ShapeInfo[] = JSON.parse(content);
    shapes.push(...categoryShapes);
  }

  return shapes;
}

/**
 * Generate PowerPoint file for a shape
 */
async function generateShapePptx(shape: ShapeInfo): Promise<Buffer> {
  const pres = new pptxgen();

  // Set presentation properties
  pres.layout = "LAYOUT_16x9";
  pres.author = "Shape Library Preview Generator";

  // Add a slide with white background
  const slide = pres.addSlide();

  // Add the shape
  const shapeDef = shape.pptxDefinition;
  slide.addShape(shapeDef.type, {
    x: shapeDef.x,
    y: shapeDef.y,
    w: shapeDef.w,
    h: shapeDef.h,
    fill: shapeDef.fill,
    line: shapeDef.line,
    shadow: shapeDef.shadow,
    rotate: shapeDef.rotate,
    flipH: shapeDef.flipH,
    flipV: shapeDef.flipV,
  });

  // Generate as buffer
  const data = await pres.write({ outputType: "nodebuffer" }) as Buffer;
  return data;
}

/**
 * Save shape as PowerPoint file
 * (In a full implementation, you would then convert this to PNG)
 */
async function generatePreviewForShape(shape: ShapeInfo): Promise<void> {
  try {
    console.log(`Generating preview for: ${shape.name} (${shape.id})`);

    // Generate PowerPoint
    const pptxData = await generateShapePptx(shape);

    // For now, save as .pptx file
    // In a full implementation, you would use a library or PowerPoint COM API
    // to export this as a PNG image
    const categoryDir = join(ASSETS_DIR, shape.category);
    if (!existsSync(categoryDir)) {
      mkdirSync(categoryDir, { recursive: true });
    }

    const pptxPath = join(categoryDir, `${shape.id}.pptx`);
    writeFileSync(pptxPath, pptxData);

    console.log(`  ✓ Saved PowerPoint: ${pptxPath}`);
    console.log(`  ⚠ Manual step: Open ${shape.id}.pptx and export as ${shape.preview}`);

  } catch (error) {
    console.error(`  ✗ Failed to generate preview for ${shape.name}:`, error);
  }
}

/**
 * Main function
 */
async function main() {
  console.log("=== Shape Preview Generator ===\n");

  // Ensure assets directory exists
  if (!existsSync(ASSETS_DIR)) {
    mkdirSync(ASSETS_DIR, { recursive: true });
  }

  // Load all shapes
  console.log("Loading shapes...");
  const shapes = loadAllShapes();
  console.log(`Found ${shapes.length} shapes across ${new Set(shapes.map(s => s.category)).size} categories\n`);

  // Generate previews
  console.log("Generating preview PowerPoint files...\n");

  for (const shape of shapes) {
    await generatePreviewForShape(shape);
  }

  console.log("\n=== Generation Complete ===");
  console.log("\nNext steps:");
  console.log("1. Open each .pptx file in the assets/ directory");
  console.log("2. Export each slide as PNG:");
  console.log("   - File → Export → Change File Type → PNG");
  console.log("   - Or use File → Save As → PNG");
  console.log("3. Save with the filename specified in the JSON (e.g., 'arrow-right.png')");
  console.log("4. Delete the .pptx files after exporting\n");

  console.log("Alternatively, on Windows, you can use PowerPoint COM automation");
  console.log("to export slides automatically. See PREVIEW_GENERATION.md for details.");
}

// Run the script
main().catch(console.error);
