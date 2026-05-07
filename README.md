# EEE202 Project Evaluator

Automated grading tool for student project submissions in EEE202 Signals and Systems courses. Designed for batch processing of PDF submissions downloaded from LMS systems.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python: 3.8+](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![DOI](https://zenodo.org/badge/1231701087.svg)](https://doi.org/10.5281/zenodo.20066611)

## Overview

This tool evaluates student PDFs against a fixed rubric for a Signals and Systems project covering continuous and discrete-time convolution, FFT analysis, and comparative evaluation. It scores submissions on a 100-point scale across five sections (1A, 1B, 2A, 2B, Conclusion) and generates a detailed Excel report.

The grader was built around a specific assignment used in EEE202 at Istinye University. Other instructors can adapt the rubric, parameters, and keyword sets in the source file for their own assignments.

## Key Features

- **Section-aware scoring** of code, graphs, and discussions across five rubric sections
- **Template-aware extraction** that ignores instructional pages students forget to delete
- **Identity from PDF content** rather than the filename, since LMS systems often rename files on download
- **Automatic file renaming** to `{StudentNo}_{NameSurname}_Project.pdf` for instructor convenience
- **Visual analysis** of pages to detect graphs, screenshots, and duplicated figures (PyMuPDF + OpenCV)
- **Bilingual support** (Turkish + English headings)
- **Detailed Excel report** with per-criterion scores, confidence values, and reasoning text

## Installation

Requires Python 3.8 or later.

```bash
pip install -r requirements.txt
```

Or directly:

```bash
pip install pdfplumber pymupdf opencv-python-headless pillow openpyxl numpy
```

## Quick Start

```bash
# Place all student PDFs in a folder, then:
python3 degerlendir_v6.py /path/to/student/pdfs

# Or run from the script's own folder:
python3 degerlendir_v6.py
```

The script writes `notlar_v6.xlsx` to the same folder. Files are also renamed in place to canonical format if `RENAME_FILES = True` (default).

> **First run:** Back up the folder before the first batch run, since file renaming is in-place.

## Configuration

Two flags at the top of `degerlendir_v6.py`:

```python
RENAME_FILES = True       # Auto-rename files to {StudentNo}_{Name}_Project.pdf
PENALIZE_FILENAME = False # If True: -5 for non-standard filenames (legacy)
```

The legacy behavior penalizes filenames that don't follow the assignment's naming rule. With LMS workflows where the system assigns its own filenames, this penalty is unfair, so the default is off.

## Output

The Excel file contains five sheets:

| Sheet | Contents |
|---|---|
| **Notlar** | Per-student scores, totals, penalties, notes |
| **Detayli Puanlar** | All 30+ criteria scored individually |
| **Kanit** | Per-criterion confidence and reasoning text |
| **Beklenen Degerler** | Reference values for all N=0..99 |
| **Gorsel Kontrol** | Visual analysis details per page |

Each row in `Notlar` shows: student number, name, parameters (a, f₀, T), section subtotals, raw total, penalties, net score (out of 100), confidence, and warning flags.

## Penalty System

| Penalty | Default | Meaning |
|---|---|---|
| `ceza_dosya_adi` | 0 (off) | Wrong filename format (legacy, see config) |
| `ceza_etiket` | -1 | Three or more graphs missing axis labels |
| `ceza_sablon` | -5 | Student left assignment instruction pages in the PDF |

## Adapting for Your Own Assignment

This tool is purpose-built for one specific assignment. To adapt:

1. **Parameters**: edit `compute_params()` if your formula for `a`, `f0`, `T` differs.
2. **Expected values**: `compute_expected()` simulates the reference solution; replace the math with your own.
3. **Section headings**: `HEADING_PATTERNS` matches headings like `1A`, `1B`, etc. Add or modify patterns for your assignment structure.
4. **Rubric criteria**: each `evaluate_*` function defines what to look for in code/graphs/discussion. Modify keyword lists, regex, and score weights.
5. **Template detection**: `TEMPLATE_INSTRUCTION_MARKERS` lists phrases that appear only in the assignment instructions. Replace with phrases from your own template.

## Limitations

- Rubric is hard-coded for the specific Signals and Systems project (continuous/discrete convolution + FFT)
- Discussion scoring relies on keyword matching and is conservative; manual review is recommended for borderline submissions (the `Manuel?` column in the Excel report flags these)
- Visual analysis can produce false positives on table-heavy pages

## License

MIT License. See [LICENSE](LICENSE) for full text.

## Citation

If you use this tool in academic work, please cite:

```bibtex
@software{dikmen_eee202_evaluator_2026,
  author       = {Dikmen, İsmail Can},
  title        = {EEE202 Project Evaluator},
  version      = {6.0},
  year         = {2026},
  publisher    = {Istinye University},
  url          = {https://github.com/DrCanD/eee202-project-evaluator}
}
```

## Author

**İ. Can Dikmen**
Istinye University, Department of Electrical and Electronics Engineering
[can.dikmen@istinye.edu.tr](mailto:can.dikmen@istinye.edu.tr)
ORCID: [0000-0002-7747-7777](https://orcid.org/0000-0002-7747-7777)

## Contributing

Issues and pull requests are welcome. For substantial changes, please open an issue first to discuss the proposed change.

---

## Hızlı Başlangıç (Türkçe)

EEE202 Sinyaller ve Sistemler dersinde verilen sürekli/ayrık konvolüsyon + FFT projesinin otomatik değerlendirme aracıdır. Ödev PDF'lerini bir klasöre koyun, scripti çalıştırın, `notlar_v6.xlsx` raporunu alın.

```bash
pip install -r requirements.txt
python3 degerlendir_v6.py /odev/pdfler
```

Önemli notlar:
- LMS'den indirilen dosyalar sistemin verdiği isimle gelir; script kimliği PDF içinden çıkarır ve dosyayı `{ÖğrenciNo}_{AdSoyad}_Project.pdf` formatına çevirir.
- Şablonun talimat sayfalarını silmeden teslim eden öğrencilere -5 ceza uygulanır (`ceza_sablon`).
- Dosya adı cezası varsayılan olarak kapalıdır (LMS dosya adlandırması nedeniyle).
- İlk çalıştırmadan önce klasörün yedeğini almanız önerilir.
