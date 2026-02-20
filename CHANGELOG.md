# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-02-14

### Added
- Multi-table selection with persistent checked state across filter operations
- Solution-aware filtering for tables and relationships
- Selected Tables panel showing all checked tables regardless of filter
- Nested tab structure (Table tabs → Relationship type subtabs)
- Excel export with three options:
  - Export Selected Table(s) - Exports all checked tables
  - Export All Tables - Exports all loaded tables
- Smart Excel sheet naming:
  - Format: Display Name (logical_name)
  - Automatic sanitization of invalid characters
  - Smart truncation for 31-character Excel limit
  - Duplicate name handling with counters
- Comprehensive relationship metadata export:
  - 1:N, N:1, and N:N relationship types
  - Complete cascade behaviors (Delete, Assign, Share, Unshare, Reparent, Merge, Rollup View)
  - Is Customizable and Is Managed flags
- Filter textbox for quick table search
- Check All / Uncheck All Selected bulk operations
- About dialog with developer contact information
- Settings persistence:
  - Last used organization
  - Last selected solution
  - Auto-open after export preference
- Comprehensive error logging:
  - Info level: Operational tracking
  - Warning level: Non-critical issues
  - Error level: Full exception details with stack traces
  - Logs written to XrmToolBox log files

### Technical
- .NET Framework 4.6.2
- XrmToolBox 1.2023.12.68 compatibility
- Microsoft.CrmSdk.CoreAssemblies 9.0.2.56
- IAboutPlugin interface implementation
- XrmToolBox-compliant logging system

### Fixed
- Excel sheet naming errors for tables with invalid characters in names
- Filter preserving checked state when searching
- BackgroundWorker progress reporting configuration
- Proper exception handling with graceful degradation (individual table export failures don't stop entire export)

## [Unreleased]

### Planned Features
- Remember last checked tables across sessions
- Default export path preference
- Auto-load tables on connection option
- Toggle between display names and logical names
- System table filtering option
- Managed relationship filtering option
- Settings UI panel for user preferences

---

For more information, see [README.md](README.md).