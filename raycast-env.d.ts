/// <reference types="@raycast/api">

/* 🚧 🚧 🚧
 * This file is auto-generated from the extension's manifest.
 * Do not modify manually. Instead, update the `package.json` file.
 * 🚧 🚧 🚧 */

/* eslint-disable @typescript-eslint/ban-types */

type ExtensionPreferences = {
  /** Enable Cache - Cache shape definitions for faster loading */
  "enableCache": boolean,
  /** Auto Cleanup Temp Files - Automatically delete temporary PowerPoint files after 60 seconds */
  "autoCleanup": boolean,
  /** Library Folder - Absolute path to store shapes JSON, previews and native PPTX. Leave empty to use app data. */
  "libraryPath"?: string,
  /** Auto-save after capture - Automatically save the captured shape to the library without showing the form */
  "autoSaveAfterCapture": boolean,
  /** Force Exact Shapes Only - Block open/copy if there is no native PPTX (100% fidelity) */
  "forceExactShapes": boolean,
  /** Use PPTX Library Deck - Store shapes inside a single PPTX deck and copy from it */
  "useLibraryDeck": boolean,
  /** Skip native PPTX save at capture - Avoid saving a PPTX during capture (faster, more reliable). Native insert still works and you can save later. */
  "skipNativeSave": boolean,
  /** Default Category - Category to show by default */
  "defaultCategory": "all" | "basic" | "arrows" | "flowchart" | "callouts"
}

/** Preferences accessible in all the extension's commands */
declare type Preferences = ExtensionPreferences

declare namespace Preferences {
  /** Preferences accessible in the `shape-picker` command */
  export type ShapePicker = ExtensionPreferences & {}
  /** Preferences accessible in the `capture-shape` command */
  export type CaptureShape = ExtensionPreferences & {}
}

declare namespace Arguments {
  /** Arguments passed to the `shape-picker` command */
  export type ShapePicker = {}
  /** Arguments passed to the `capture-shape` command */
  export type CaptureShape = {}
}

