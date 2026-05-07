# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [6.0] - 2026-05-07

### Added
- Template-aware identity extraction: the script now finds the
  `STUDENT SUBMISSION TEMPLATE` marker and only reads student information
  from after that point. Stops the script from picking up the example
  student number embedded in the assignment instructions.
- Template-cleanliness penalty (`ceza_sablon = -5`) for submissions that
  include the assignment instruction pages (1-9) instead of only the
  student-fill section.
- Automatic file renaming: after identity is detected from the PDF text,
  files are renamed in place to `{StudentNo}_{NameSurname}_Project.pdf`.
  Aksanlar ASCII'ye çevrilir.
- Configuration flags `RENAME_FILES` and `PENALIZE_FILENAME` at the top
  of the script.
- New columns and warning notes in the Excel report:
  `SABLON?`, `RENAME`, `KIMLIK`, `KIMLIK_FARK`, `DOSYA_ADI`.

### Changed
- Identity extraction now reads only from inside the PDF text. Filename
  is used only as a last-resort fallback (with a manual-review flag).
  Reason: LMS systems often rename submitted files on download, so the
  filename is not a reliable identity source.
- Filename format penalty (`ceza_dosya_adi`) is off by default. Can be
  re-enabled via `PENALIZE_FILENAME = True`.
- `split_code_and_discussion()` now matches section headings anywhere
  in a line (not just at the start), so headings like
  `1A. Convolution - MATLAB Code` are detected correctly.
- `extract_reported_params()` skips the assignment-template region,
  preventing it from reading mod-operator divisors as student parameter
  values.
- Visual analysis ignores assignment instruction pages, removing false
  positives in the `GRAFIK_TEKRAR?` (duplicate-figure) detector.

### Fixed
- Critical bug where students who included the assignment template in
  their PDF received drastically wrong scores due to:
  - Wrong student number (assignment example was picked up)
  - Wrong reported parameters (mod-operator divisors were read as values)
  - Phantom duplicate-figure warnings between instruction-page tables

## [5.x] - 2026-04-03

### Added
- PDF ligature repair (`ffi`/`ffl` ligatures broken into separate
  characters are rejoined).
- Visual graph evidence layer (PyMuPDF + OpenCV).
- Prefix-stem keyword matching (`attenuat`, `suppress`, `preserv`, ...).

### Changed
- Decimal-number sentence splitting fix.
- More robust text extraction with quality metrics and fallback chain.

## [4.x] - 2026-03-25

### Added
- Bilingual support (Turkish + English heading patterns).
- Windows compatibility (subprocess timeout, safe console output).

### Changed
- `find_section_boundaries()` rewritten to use a unified pattern table.
