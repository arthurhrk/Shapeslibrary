/**
 * Save captured shapes to JSON files
 */

import { readFileSync, writeFileSync, existsSync, mkdirSync } from "fs";
import { join } from "path";
import { environment } from "@raycast/api";
import { getShapesDir as getShapesDirUtil } from "./paths";
import { ShapeInfo, ShapeCategory } from "../types/shapes";

/**
 * Get path to shapes directory
 */
function getShapesDir(): string {
  return getShapesDirUtil();
}

/**
 * Get path to category JSON file
 */
function getCategoryFilePath(category: ShapeCategory): string {
  const shapesDir = getShapesDir();
  return join(shapesDir, `${category}.json`);
}

/**
 * Load existing shapes from a category file
 */
function loadCategoryShapes(category: ShapeCategory): ShapeInfo[] {
  const filePath = getCategoryFilePath(category);

  if (!existsSync(filePath)) {
    return [];
  }

  try {
    const content = readFileSync(filePath, "utf-8");
    return JSON.parse(content);
  } catch (error) {
    console.error(`Failed to load ${category} shapes:`, error);
    return [];
  }
}

/**
 * Save shapes to a category file
 */
function saveCategoryShapes(category: ShapeCategory, shapes: ShapeInfo[]): void {
  const filePath = getCategoryFilePath(category);

  try {
    // Sort shapes alphabetically by name
    const sortedShapes = shapes.sort((a, b) => a.name.localeCompare(b.name));

    // Write with pretty formatting
    const json = JSON.stringify(sortedShapes, null, 2);
    writeFileSync(filePath, json, "utf-8");
    try {
      console.log(`[ShapeSaver] Saved ${category}.json to ${filePath}`);
    } catch {}
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    throw new Error(`Failed to save shapes to ${category}.json at ${filePath}: ${msg}`);
  }
}

/**
 * Check if a shape with the same ID already exists
 */
export function shapeExists(id: string, category: ShapeCategory): boolean {
  const shapes = loadCategoryShapes(category);
  return shapes.some((shape) => shape.id === id);
}

/**
 * Add a captured shape to the library
 */
export function addShapeToLibrary(shape: ShapeInfo): string {
  const { category } = shape;

  // Load existing shapes
  const shapes = loadCategoryShapes(category);

  // Check if shape ID already exists
  const existingIndex = shapes.findIndex((s) => s.id === shape.id);

  if (existingIndex !== -1) {
    // Replace existing shape
    shapes[existingIndex] = shape;
  } else {
    // Add new shape
    shapes.push(shape);
  }

  // Save updated shapes
  saveCategoryShapes(category, shapes);
  return getCategoryFilePath(category);
}

/**
 * Update an existing shape in the library
 */
export function updateShapeInLibrary(id: string, category: ShapeCategory, updates: Partial<ShapeInfo>): void {
  const shapes = loadCategoryShapes(category);

  const index = shapes.findIndex((s) => s.id === id);

  if (index === -1) {
    throw new Error(`Shape with ID '${id}' not found in ${category} category`);
  }

  // Update shape
  shapes[index] = {
    ...shapes[index],
    ...updates,
    id, // Ensure ID doesn't change
    category, // Ensure category doesn't change
  };

  saveCategoryShapes(category, shapes);
}

/**
 * Remove a shape from the library
 */
export function removeShapeFromLibrary(id: string, category: ShapeCategory): void {
  const shapes = loadCategoryShapes(category);

  const filteredShapes = shapes.filter((s) => s.id !== id);

  if (filteredShapes.length === shapes.length) {
    throw new Error(`Shape with ID '${id}' not found in ${category} category`);
  }

  saveCategoryShapes(category, filteredShapes);
}

/**
 * Get count of shapes in each category
 */
export function getShapeCounts(): Record<ShapeCategory, number> {
  const categories: ShapeCategory[] = ["basic", "arrows", "flowchart", "callouts"];

  const counts: Record<string, number> = {};

  categories.forEach((category) => {
    const shapes = loadCategoryShapes(category);
    counts[category] = shapes.length;
  });

  return counts as Record<ShapeCategory, number>;
}

/**
 * Get total number of shapes across all categories
 */
export function getTotalShapeCount(): number {
  const counts = getShapeCounts();
  return Object.values(counts).reduce((sum, count) => sum + count, 0);
}
