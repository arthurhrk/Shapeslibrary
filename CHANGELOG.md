# Changelog

All notable changes to the PowerPoint Shapes Library extension will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2025-01-28

### Added
- **Repair Broken Previews**: New manual action to fix orphaned preview thumbnails
  - Accessible via keyboard shortcut `Cmd/Ctrl + Shift + R`
  - Automatically finds and moves preview files to correct category folders
  - Shows count of repaired previews with success message
  - Located in Utility section of action menu
- Auto-repair function that scans and fixes misplaced preview files
  - Runs on-demand when user invokes "Repair Broken Previews"
  - Creates marker file to track repair status

### Changed
- **Improved Category Display**: Category names now use "Capitalize Each Word" format instead of ALL CAPS
  - "BASIC SHAPES" → "Basic Shapes"
  - "PROPOSALS" → "Proposals"
  - "VISUALS" → "Visuals"
  - "LEGAL" → "Legal"
  - "NATIVE-ONLY" → "Native-Only"
- Enhanced category change workflow to physically move preview PNG files
  - When editing a shape and changing its category, the preview file is now automatically moved to the new category folder
  - Ensures thumbnails remain correctly linked after category changes

### Fixed
- **Thumbnail Display Bug**: Fixed issue where thumbnails would show as blue squares after changing shape category
  - Preview files are now properly moved when shapes are reassigned to different categories
  - Prevents broken preview paths in shape metadata
- Added fallback copy mechanism when file rename fails across different drives/devices

## [1.0.0] - 2025-01-XX

### Added
- Initial release of PowerPoint Shapes Library
- Visual grid browser with PNG/SVG preview system
- Shape capture from PowerPoint using COM API (Windows)
- Customizable category system with editable names
- Edit and delete shapes functionality
- Multiple shape insertion methods:
  - Direct insertion into active PowerPoint (Windows)
  - Copy to clipboard
  - Open in new PowerPoint file
- Shape library import/export via ZIP files
- Batch preview generation script (Windows only)
- Smart caching system for improved performance
- PPTX library deck for faster shape access
- Auto-cleanup of temporary files
- Cross-platform support (Windows full features, macOS basic support)
- TypeScript implementation with full type safety
- Keyboard shortcuts for all major actions
- Configurable preferences for workflow customization

[1.1.0]: https://github.com/yourusername/shapes-library/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/yourusername/shapes-library/releases/tag/v1.0.0
