#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportOptionalMemberAccess=false, reportAttributeAccessIssue=false
"""
EEE202 Sinyaller ve Sistemler — Otomatik Proje Degerlendirme Sistemi
====================================================================

Ogrenci proje PDF'lerini toplu olarak puanlayan, sablon temizligi/kimlik
denetimi yapan ve sonuclari Excel'e raporlayan acik kaynak bir
degerlendirme aracidir.

Ana ozellikler
--------------
- Bolum bazli puanlama (1A/1B/2A/2B + Sonuc, 100 puan uzerinden)
- Sablon talimat sayfalarinin PDF'e dahil edilip edilmediginin tespiti
- LMS-renamed dosyalarda kimligin PDF iceriginden cikarilmasi
- Tespit sonrasi dosyalarin {StudentNo}_{NameSurname}_Project.pdf formatina
  otomatik adlandirilmasi
- Grafik/screenshot/kopya sayfalarin gorsel analizi (PyMuPDF + OpenCV)
- Cok-katmanli metin cikarma (fitz + pdfplumber fallback)
- 5 sayfali Excel raporu (kanit ve gerekce sutunlariyla)

Yazar       : I. Can Dikmen
Kurum       : Istinye Universitesi, Elektrik Elektronik Muhendisligi Bolumu
E-posta     : can.dikmen@istinye.edu.tr
ORCID       : 0000-0002-7747-7777
Repo        : https://github.com/DrCanD/eee202-project-evaluator
Lisans      : MIT
Versiyon    : 6.0
Olusturulma : 2026-03-25
Son guncelleme : 2026-05-07

Kullanim
--------
    python3 degerlendir_v6.py                  # script klasorunu tara
    python3 degerlendir_v6.py /pdf/klasoru     # belirtilen klasoru tara

Bagimliliklar
-------------
    pip install pdfplumber pymupdf opencv-python-headless pillow openpyxl numpy

Onemli notlar
-------------
RENAME_FILES varsayilan olarak True'dur; ilk kosumda yedek almaniz onerilir.
PENALIZE_FILENAME varsayilan False'tur cunku LMS dosyalari sistemin verdigi
adla geldigi icin dosya adi cezasi adil olmaz; eski davranisi istiyorsaniz
True yapabilirsiniz.

Atif
----
Bu araci akademik bir calismada kullanirsaniz lutfen su sekilde atif yapiniz:

    Dikmen, I. C. (2026). EEE202 Otomatik Proje Degerlendirme Sistemi
    (Surum 6.0) [Bilgisayar yazilimi]. Istinye Universitesi.
    https://github.com/DrCanD/eee202-project-evaluator

Surum gecmisi
-------------
6.0 (2026-05-07) - Sablon-bilincli kimlik cikarma, otomatik dosya adlandirma,
                   sablon temizligi cezasi (-5), LMS uyumluluk modu.
5.x (2026-04-03) - PDF ligature onarimi, gorsel kanit, prefix-stem keyword
                   eslestirme.
4.x (2026-03-25) - TR + EN iki dilli destek, Windows uyumlulugu.
"""

import os
import re
import sys
import math
import glob
import argparse
import subprocess
import tempfile
import unicodedata
from dataclasses import dataclass
from statistics import median
from typing import Any, Dict, List, Tuple, Optional

import numpy as np

try:
    import pdfplumber
except ImportError:
    print("HATA: pdfplumber yuklu degil.\n  pip install pdfplumber")
    sys.exit(1)

try:
    import fitz
except ImportError:
    print("HATA: PyMuPDF yuklu degil.\n  pip install pymupdf")
    sys.exit(1)

try:
    import cv2
except ImportError:
    print("HATA: opencv-python yuklu degil.\n  pip install opencv-python")
    sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("HATA: Pillow yuklu degil.\n  pip install pillow")
    sys.exit(1)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.worksheet import Worksheet
except ImportError:
    print("HATA: openpyxl yuklu degil.\n  pip install openpyxl")
    sys.exit(1)


# ---------------------------------------------------------------------------
# KONFIGURASYON (kullanici degistirebilir)
# ---------------------------------------------------------------------------

# LMS'den indirilen PDF'ler genellikle sistemin verdigi adla geliyor; bu yuzden
# kimlik DOSYA ADINDAN DEGIL, PDF iceriginden cikariliyor. Asagidaki bayraklar
# eski davranisi geri getirmek icin kullanilabilir.

# Identity tespit edildikten sonra dosyayi {ogr_no}_{Ad}_Project.pdf formatina
# getir. Hocanin sonradan ogrenciyi bulmasini kolaylastirir.
RENAME_FILES = True

# True yaparsan eski davranis: dosya adi formati yanlissa -5 ceza. LMS rename
# yapiyorsa bu cezayi 0'a (False) birakmak adildir.
PENALIZE_FILENAME = False


# ---------------------------------------------------------------------------
# TEMEL HESAPLAR
# ---------------------------------------------------------------------------

def compute_params(N: int) -> Dict[str, int]:
    return {
        "a": 1 + (N % 5),
        "f0": 2 + (N % 7),
        "T": 3 + (N % 4),
    }


def compute_expected(a: int, f0: int, T: int) -> Dict[str, float]:
    dt = 0.001
    Ts = 0.01

    t_x = np.arange(0, T, dt)
    x = np.sin(2 * np.pi * f0 * t_x)
    t_h = np.arange(0, 5.0 / a, dt)
    h = np.exp(-a * t_h)
    y = np.convolve(x, h) * dt
    y_peak = float(np.max(y))
    y_peak_t = float(np.argmax(y) * dt)
    y_duration = float((len(y) - 1) * dt)

    Fs = 1 / dt
    N_fft = len(x)
    X_fft = np.abs(np.fft.fft(x)) / N_fft
    f_axis = np.arange(N_fft) * Fs / N_fft
    half = N_fft // 2
    dominant_f = float(f_axis[1 + np.argmax(X_fft[1:half])])
    H_f0 = 1.0 / math.sqrt(a**2 + (2 * math.pi * f0) ** 2)
    H_0 = 1.0 / a
    freq_3db = a / (2 * math.pi)
    attenuation_pct = (1 - H_f0 / H_0) * 100

    Fs_d = 1 / Ts
    n_x = np.arange(0, int(T / Ts))
    x_d = np.sin(2 * np.pi * f0 * n_x * Ts)
    n_h = np.arange(0, int(5.0 / (a * Ts)) + 1)
    h_d = np.exp(-a * n_h * Ts)
    y_d = np.convolve(x_d, h_d)
    y_d_peak = float(np.max(y_d))
    y_d_peak_scaled = y_d_peak * Ts
    y_d_duration = float((len(y_d) - 1) * Ts)

    df_cont = Fs / len(x)
    df_disc = Fs_d / len(x_d)
    nyquist_ok = Fs_d > 2 * f0
    alias_if_ts_0_1 = (10 > 2 * f0) is False

    return {
        "dt": dt,
        "Ts": Ts,
        "Fs": Fs,
        "Fs_d": Fs_d,
        "y_peak": round(y_peak, 4),
        "y_peak_t": round(y_peak_t, 4),
        "y_duration": round(y_duration, 4),
        "dominant_f": round(dominant_f, 2),
        "H_f0": round(H_f0, 6),
        "H_0": round(H_0, 4),
        "freq_3db": round(freq_3db, 4),
        "attenuation_pct": round(attenuation_pct, 1),
        "y_d_peak": round(y_d_peak, 4),
        "y_d_peak_scaled": round(y_d_peak_scaled, 4),
        "y_d_duration": round(y_d_duration, 4),
        "df_cont": round(df_cont, 4),
        "df_disc": round(df_disc, 4),
        "nyquist_ok": nyquist_ok,
        "alias_if_ts_0_1": alias_if_ts_0_1,
        "two_f0": round(2 * f0, 4),
    }


# ---------------------------------------------------------------------------
# METIN / YAPI ISLEME
# ---------------------------------------------------------------------------

PDF_TEXT_TIMEOUT_SEC = 12


def _clean_extracted_text(text: Optional[str]) -> str:
    if not text:
        return ""
    cleaned = []
    for ch in text:
        if ch in "\n\r\t" or ord(ch) >= 32:
            cleaned.append(ch)
        else:
            cleaned.append(" ")
    text = "".join(cleaned).replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _text_quality_metrics(text: Optional[str]) -> Dict[str, float]:
    if not text:
        return {"non_ws_len": 0, "word_count": 0, "avg_word_len": 0.0,
                "vowel_ratio": 0.0, "punct_ratio": 1.0, "cid_count": 0, "short_word_ratio": 1.0}
    words = re.findall(r"[A-Za-zÇçĞğİıÖöŞşÜü']+", text)
    total_letters = sum(len(w) for w in words)
    vowel_count = sum(sum(ch.lower() in "aeiouöüıi" for ch in w) for w in words)
    short_count = sum(len(w) <= 2 for w in words)
    non_ws = [ch for ch in text if not ch.isspace()]
    return {
        "non_ws_len": len(non_ws),
        "word_count": len(words),
        "avg_word_len": (total_letters / len(words)) if words else 0.0,
        "vowel_ratio": (vowel_count / total_letters) if total_letters else 0.0,
        "punct_ratio": sum((not ch.isalnum()) and not ch.isspace() for ch in text) / max(1, len(text)),
        "cid_count": len(re.findall(r"\(cid:\d+\)", text)),
        "short_word_ratio": (short_count / len(words)) if words else 1.0,
    }


def _looks_readable_text(text: Optional[str]) -> bool:
    m = _text_quality_metrics(text)
    if m["non_ws_len"] < 40:
        return False
    if m["cid_count"] >= 8:
        return False
    if m["punct_ratio"] > 0.35:
        return False
    if m["word_count"] >= 50 and m["avg_word_len"] < 3.05:
        return False
    if m["word_count"] >= 50 and m["short_word_ratio"] > 0.58:
        return False
    if m["word_count"] >= 50 and m["vowel_ratio"] < 0.23:
        return False
    if m["word_count"] >= 20:
        return True
    return (m["non_ws_len"] >= 80 and m["avg_word_len"] >= 4.0
            and m["punct_ratio"] < 0.25 and m["vowel_ratio"] >= 0.25)


def _extract_text_fitz(pdf_path: str) -> Optional[str]:
    try:
        doc = fitz.open(pdf_path)
        try:
            pages = []
            for idx in range(doc.page_count):
                page = doc.load_page(idx)
                pt = page.get_text("text", sort=True)
                pages.append(pt if isinstance(pt, str) else "")
        finally:
            doc.close()
        return _clean_extracted_text("\n".join(pages))
    except Exception:
        return None


def _extract_text_pdfplumber_direct(pdf_path: str) -> Optional[str]:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pages = []
            for page in pdf.pages:
                try:
                    pages.append(page.extract_text() or "")
                except Exception:
                    pages.append("")
        return _clean_extracted_text("\n".join(pages))
    except Exception:
        return None


def _extract_text_pdfplumber_with_timeout(pdf_path: str, timeout_sec: int = PDF_TEXT_TIMEOUT_SEC) -> Optional[str]:
    worker_path = None
    try:
        with tempfile.NamedTemporaryFile("w", encoding="utf-8", delete=False, suffix=".pdftext.tmp") as tmp:
            worker_path = tmp.name
        proc = subprocess.Popen(
            [sys.executable, os.path.abspath(__file__), "--extract-text-worker", pdf_path, worker_path],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        try:
            proc.wait(timeout=timeout_sec)
        except (subprocess.TimeoutExpired, KeyboardInterrupt):
            try:
                proc.kill()
            except OSError:
                pass
            try:
                proc.wait(timeout=3)
            except Exception:
                pass
            return None
        if proc.returncode != 0 or not os.path.exists(worker_path):
            return None
        with open(worker_path, "r", encoding="utf-8") as fh:
            text = fh.read().strip()
        return text or None
    except Exception:
        return None
    finally:
        if worker_path and os.path.exists(worker_path):
            try:
                os.remove(worker_path)
            except OSError:
                pass


def extract_text(pdf_path: str) -> Optional[str]:
    """Try fitz first (fast), fall back to pdfplumber if quality is poor."""
    fitz_text = _extract_text_fitz(pdf_path)
    if _looks_readable_text(fitz_text):
        return fitz_text
    # Subprocess-based pdfplumber (with timeout protection)
    try:
        plumber_text = _extract_text_pdfplumber_with_timeout(pdf_path)
        if _looks_readable_text(plumber_text):
            return plumber_text
    except Exception:
        pass
    # Last resort: direct pdfplumber (no timeout, for systems where subprocess fails)
    try:
        direct_text = _extract_text_pdfplumber_direct(pdf_path)
        if _looks_readable_text(direct_text):
            return direct_text
    except Exception:
        pass
    return fitz_text if _looks_readable_text(fitz_text) else None


def _run_extract_text_worker(pdf_path: str, output_path: str) -> int:
    text = _extract_text_pdfplumber_direct(pdf_path) or ""
    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write(text)
    return 0


def strip_accents(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def normalize_text(text: str) -> str:
    if text is None:
        return ""
    text = strip_accents(text)
    text = text.lower()
    # PDF ligature fix: collapse common broken function names
    # ffi/ffl ligatures often break into separate chars in PDF extraction
    text = re.sub(r"\bf\s*f\s*t\b", "fft", text)       # f f t → fft
    text = re.sub(r"\bc\s*o\s*n\s*v\b", "conv", text)   # c o n v → conv
    # Hyphen between letters → space (low-frequency → low frequency)
    # but keep exp(-a), -3dB, negative signs intact
    text = re.sub(r"(?<=[a-z])-(?=[a-z])", " ", text)
    text = re.sub(r"[\t\r ]+", " ", text)
    return text


def normalize_name(name: str) -> str:
    name = re.sub(r"[_-]+", " ", name or "")
    name = re.sub(r"\s+", " ", name).strip()
    return name


def safe_console_text(value: object) -> str:
    """Encode text safely for console output (handles Turkish chars on Windows)."""
    text = unicodedata.normalize("NFC", str(value))
    encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
    return text.encode(encoding, errors="replace").decode(encoding, errors="replace")


def normalize_student_number(value: str) -> Optional[str]:
    """Extract digits from a potentially space-separated student number."""
    digits = re.sub(r"\D", "", value or "")
    return digits if 8 <= len(digits) <= 12 else None


def clean_name(value: str) -> str:
    value = value or ""
    value = re.split(r"(?:ogrenci\s*numarasi|ogrenci\s*no|ders\s*adi|tarih|konu)\b", value, flags=re.I)[0]
    value = re.sub(r"[^A-Za-zÇçĞğİıÖöŞşÜü\s_-]", " ", value)
    value = re.sub(r"[_-]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _find_student_section_start(lines: List[str]) -> int:
    """Sablonun talimat sayfalari PDF'e dahil edildiginde, ogrenci girdi
    alaninin baslangicini bul. Bulunamazsa 0 dondur (eskiden gibi tum metni
    tara). 'STUDENT SUBMISSION TEMPLATE' / 'A. Student Information' /
    'B. Parameter Table' veya Turkce muadilleri."""
    for i, line in enumerate(lines):
        nline = normalize_text(line).strip()
        if (re.search(r"student\s+submission\s+template", nline)
                or re.search(r"ogrenci\s+gonderim", nline)
                or re.search(r"^a\.\s*(student\s+information|ogrenci\s+bilgileri)", nline)
                or re.search(r"^a\)\s*(student\s+information|ogrenci\s+bilgileri)", nline)
                or re.search(r"^b\.\s*(parameter\s+table|parametre\s+tablo)", nline)):
            return i
    return 0


# Sablonun talimat sayfalarinda gecen ve ogrenci alaninda bulunmayan
# karakteristik ifadeler. Bunlardan yeterli sayida (>=3) varsa, ogrenci
# odev sablonunu silmeden (talimat sayfalariyla birlikte) gondermis demektir.
TEMPLATE_INSTRUCTION_MARKERS = [
    "project objective and scope",
    "submission format and mandatory rules",
    "common student mistakes",
    "quick reference checklist",
    "if the student number is 2021030547",
    "1. project objective",
    "2. submission format",
    "3. parameter determination",
    "4. mandatory report structure",
    "8. common student mistakes",
    "5. part 1: continuous time analysis",
    "6. part 2: discrete time analysis",
    "how to write the conclusion section",
]


def detect_template_uncleaned(text: str) -> Tuple[bool, List[str]]:
    """Sablonun (1-9. sayfa) talimat sayfalarinin PDF'e dahil edilip
    edilmedigini tespit et. >=3 marker varsa True dondur."""
    nt = normalize_text(text)
    found = [m for m in TEMPLATE_INSTRUCTION_MARKERS if m in nt]
    return (len(found) >= 3, found)


def make_filename_safe_name(name: str) -> str:
    """Ad-soyadi dosya adi guvenli hale getir.
    Ornek: 'Kübra ALKAN' -> 'KubraAlkan', 'Mert Akman' -> 'MertAkman'.

    - Aksanlar/Turkce ozel karakterler ASCII'ye cevrilir (cross-platform guvenli)
    - Kelimeler title-case yapilip birlestirilir (mevcut konvansiyon)
    - Bos/gecersiz girdide bos string doner
    """
    if not name:
        return ""
    # Aksanlari sok (ş -> s, Ü -> U, vs.)
    n = strip_accents(name)
    # Sadece harf ve bosluklari tut
    n = re.sub(r"[^A-Za-z\s]", " ", n)
    parts = [p for p in re.split(r"\s+", n) if p]
    if not parts:
        return ""
    # Title case
    parts = [p[0].upper() + p[1:].lower() for p in parts]
    return "".join(parts)


def expected_canonical_filename(ogr_no: str, ogr_name: str) -> str:
    """{StudentNo}_{NameSurname}_Project.pdf formatinda istenen dosya adini uret."""
    safe_name = make_filename_safe_name(ogr_name) or "Unknown"
    return f"{ogr_no}_{safe_name}_Project.pdf"


def rename_pdf_to_canonical(pdf_path: str, ogr_no: str, ogr_name: str) -> Tuple[str, Optional[str]]:
    """PDF'i {StudentNo}_{NameSurname}_Project.pdf formatina yeniden adlandir.

    Donus: (yeni_path, hata_mesaji_ya_da_None)
    - Hedef dosya adi mevcutsayfada zaten varsa ve bu dosya degilse, _2/_3
      gibi suffix eklenir.
    - Yeniden adlandirilamadigi durumda orijinal yol ve hata mesaji doner.
    """
    if not ogr_no or not ogr_name:
        return pdf_path, "kimlik eksik"
    target_name = expected_canonical_filename(ogr_no, ogr_name)
    folder = os.path.dirname(os.path.abspath(pdf_path))
    target_path = os.path.join(folder, target_name)
    # Zaten dogru isimdeyse hicbir sey yapma
    if os.path.abspath(pdf_path) == target_path:
        return pdf_path, None
    # Hedef varsa ve farkli bir dosyaysa, suffix ekle
    if os.path.exists(target_path):
        base, ext = os.path.splitext(target_name)
        for suffix in range(2, 100):
            candidate = os.path.join(folder, f"{base}_{suffix}{ext}")
            if not os.path.exists(candidate):
                target_path = candidate
                break
        else:
            return pdf_path, "hedef ad icin yer bulunamadi (1-99)"
    try:
        os.rename(pdf_path, target_path)
        return target_path, None
    except OSError as e:
        return pdf_path, f"OS hatasi: {e}"


def extract_student_info(text: str, filename: str) -> Tuple[Optional[str], str, Optional[int], Dict[str, Any]]:
    """Ogrenci numarasi, ad, N degerini metinden cikar. Dosya adini sadece
    fallback ve capraz kontrol icin kullan.

    Donus: (ogr_no, ogr_name, N, info)
        info = {
            "src_no": "text" | "filename" | None,    # numara nereden geldi
            "src_name": "text" | "filename" | None,
            "txt_no": ...,    # metinden bulunan (yoksa None)
            "fn_no":  ...,    # dosya adindan bulunan (yoksa None)
            "fn_ok":  bool,   # dosya adi tam formatta mi
            "fn_text_match": bool,  # metin ve dosya adi numarasi uyusuyor mu
        }
    """
    fn_parts = re.match(r"(\d{8,12})_([A-Za-zÇçĞğİıÖöŞşÜü_-]+?)_(?:Proje|Project)\.pdf$", filename, re.I)
    fn_no = fn_parts.group(1) if fn_parts else None
    fn_name = normalize_name(fn_parts.group(2)) if fn_parts else None
    fn_ok = bool(fn_parts)

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    student_section_start = _find_student_section_start(lines)
    chunk_lines = lines[student_section_start:student_section_start + 80]
    first_chunk = "\n".join(chunk_lines)

    txt_no = None
    for pat in [
        r"[Oo]grenci\s*(?:[Nn]o|[Nn]umarasi)\s*[:=]?\s*([0-9][0-9\s-]{6,20}\d)",
        r"[Ss]tudent\s*(?:[Nn]o|[Nn]umber)\s*[:=]?\s*([0-9][0-9\s-]{6,20}\d)",
        r"[Nn]umara\s*[:=]?\s*([0-9][0-9\s-]{6,20}\d)",
        r"\b(20[0-9\s-]{7,15}\d)\b",
    ]:
        m = re.search(pat, first_chunk)
        if m:
            cand = normalize_student_number(m.group(1))
            # Sablonun ornek numarasini (2021030547) yakalamamak icin ek kontrol:
            # eger ogrenci alani bulunamadiysa ve aday numara taninmis sablon
            # ornekleriyle eslesiyorsa atla.
            if cand and not (student_section_start == 0 and cand == "2021030547"):
                txt_no = cand
                break

    txt_name = None
    name_search_lines = chunk_lines[:40] if student_section_start else lines[:40]
    for i, line in enumerate(name_search_lines):
        nline = normalize_text(line).strip()
        if nline.startswith("ad soyad") or nline.startswith("isim") or nline.startswith("full name") or nline.startswith("name"):
            value = re.sub(r"^(Ad\s*Soyad|Isim|Full\s*Name|Name)\s*[:=]?\s*", "", line, flags=re.I).strip()
            value = clean_name(value)
            if not value and i + 1 < len(name_search_lines):
                value = clean_name(name_search_lines[i + 1])
            if value:
                txt_name = value
                break

    if not txt_name:
        search_text = "\n".join(chunk_lines) if student_section_start else text
        m = re.search(r"(?:Ad\s*Soyad|Isim|Full\s*Name|Name)\s*[:=]?\s*([^\n]{2,60})", search_text, re.I)
        if m:
            txt_name = clean_name(m.group(1))

    # ONCELIK: SADECE METIN. LMS dosyalari sistemin adlandirma seklinde verir,
    # bu yuzden dosya adi guvenilir bir kimlik kaynagi degil. Eger metinde
    # kimlik bulunamazsa, hocanin manuel inceleme yapmasi gerekir; o zaman
    # son care olarak dosya adina dusuyoruz (yine de kullaniciya not edilir).
    if txt_no:
        ogr_no, src_no = txt_no, "text"
    elif fn_no:
        ogr_no, src_no = fn_no, "filename_fallback"
    else:
        ogr_no, src_no = None, None

    if txt_name:
        ogr_name, src_name = txt_name, "text"
    elif fn_name:
        ogr_name, src_name = fn_name, "filename_fallback"
    else:
        ogr_name, src_name = "", None

    fn_text_match = bool(txt_no and fn_no and txt_no == fn_no)
    N = int(ogr_no[-2:]) if ogr_no else None

    info = {
        "src_no": src_no, "src_name": src_name,
        "txt_no": txt_no, "fn_no": fn_no,
        "fn_ok": fn_ok,
        "fn_text_match": fn_text_match,
    }
    return ogr_no, ogr_name, N, info


def extract_reported_params(text: str) -> Dict[str, int]:
    params: Dict[str, int] = {}
    all_lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # Sablonun formul tanimlari (orn. "a = 1 + mod(N, 5)") ve ornek
    # (orn. "Example: ... a = 3, f0 = 7 Hz") satirlari ogrencinin gercek
    # parametre bildiriminden ayrilmali. Eger ogrenci alani basligi varsa
    # oradan basla; yoksa eski davranis (tum metnin ilk 140 satiri).
    section_start = _find_student_section_start(all_lines)
    lines = all_lines[section_start:]

    for line in lines[:140]:
        nline = normalize_text(line)
        nums = [int(x) for x in re.findall(r"\d+", nline)]
        if not nums:
            continue

        if "f0" not in params:
            if "f0" in nline and "mod" in nline and len(nums) >= 4:
                params["f0"] = nums[-1]
            elif re.search(r"^f0\s*[:=]", nline):
                params["f0"] = nums[-1]

        if "a" not in params:
            if re.search(r"(^|\b)a\s*=.*mod\s*\(\s*n\s*,\s*5\s*\)", nline) and len(nums) >= 4:
                params["a"] = nums[-1]
            elif re.search(r"^a\s*[:=]", nline):
                params["a"] = nums[-1]

        if "T" not in params:
            if re.search(r"(^|\b)t\s*=.*mod\s*\(\s*n\s*,\s*4\s*\)", nline) and len(nums) >= 4:
                params["T"] = nums[-1]
            elif re.search(r"^t\s*[:=]", nline) and "tarih" not in nline and "date" not in nline:
                params["T"] = nums[-1]

    # Fallback regex'leri de ogrenci alaninin metni uzerinde calistir.
    student_text = "\n".join(lines)[:6000] if section_start else text[:6000]
    tnorm = normalize_text(student_text)
    if "a" not in params:
        m = re.search(r"\ba\s*=\s*(?:1\s*\+\s*mod\s*\(\s*n\s*,\s*5\s*\)\s*=\s*)?(\d+)", tnorm)
        if m:
            params["a"] = int(m.group(1))
    if "f0" not in params:
        m = re.search(r"f0\s*=\s*(?:2\s*\+\s*mod\s*\(\s*n\s*,\s*7\s*\)\s*=\s*)?(\d+)", tnorm)
        if m:
            params["f0"] = int(m.group(1))
    if "T" not in params:
        m = re.search(r"\bt\s*=\s*(?:3\s*\+\s*mod\s*\(\s*n\s*,\s*4\s*\)\s*=\s*)?(\d+)", tnorm)
        if m:
            params["T"] = int(m.group(1))

    return params


HEADING_PATTERNS = {
    "1a": [
        r"^1a\b.*konvol",
        r"^1a\b.*convol",
        r"^1a\b",
        r"^konvolusyon\b",
        r"^convolution\b",
        r"^part\s*1\s*a\b",
        r"^bolum\s*1\s*a\b",
        r"^section\s*1\s*a\b",
        r"^part\s*1[:\s].*(?:convol|konvol|continuous|surekli)",
    ],
    "1b": [
        r"^1b\b.*fourier",
        r"^1b\b.*spectrum",
        r"^1b\b.*spektr",
        r"^1b\b",
        r"^part\s*1\s*b\b",
        r"^bolum\s*1\s*b\b",
        r"^section\s*1\s*b\b",
    ],
    "2a": [
        r"^2a\b.*(ornek|orn[eö]kleme|ayrik|sampl|discrete)",
        r"^2a\b",
        r"^ornekleme ve ayrik konvolusyon\b",
        r"^sampling and discrete convolution\b",
        r"^part\s*2\s*a\b",
        r"^bolum\s*2\s*a\b",
        r"^section\s*2\s*a\b",
    ],
    "2b": [
        r"^2b\b.*(fourier|discrete|ayrik)",
        r"^2b\b",
        r"^discrete fourier\b",
        r"^part\s*2\s*b\b",
        r"^bolum\s*2\s*b\b",
        r"^section\s*2\s*b\b",
    ],
    "sonuc": [
        r"^sonuc ve genel karsilastirma\b",
        r"^genel karsilastirma\b",
        r"^sonuc\b",
        r"^conclusion and comparative\b",
        r"^conclusion\b",
        r"^comparative evaluation\b",
    ],
}


def line_matches_heading(line: str, patterns: List[str]) -> bool:
    nline = normalize_text(line).strip()
    return any(re.search(p, nline) for p in patterns)


def find_section_boundaries(text: str) -> Dict[str, str]:
    lines = [ln.rstrip() for ln in text.splitlines()]

    # Sablonun talimat kismi PDF'e dahil edilmisse, oradaki "5.1 - 1A: Convolution"
    # gibi alt baslik referanslarini ana 1A bolumuyle karistirmamak icin tarama
    # araligini ogrenci alaniyla sinirla. Su anki HEADING_PATTERNS '^1a\b' gibi
    # satir basi anchor'larina bagli oldugu icin sablonun "5.1 - 1A:" satirlari
    # zaten eslesmez, ama ileride pattern'ler genislerse de guvende oluruz.
    nonblank = [ln.strip() for ln in lines if ln.strip()]
    section_start_in_nonblank = _find_student_section_start(nonblank)
    if section_start_in_nonblank > 0:
        # nonblank-index'i orijinal lines-index'ine geri cevir
        target_marker = nonblank[section_start_in_nonblank]
        offset = 0
        for i, ln in enumerate(lines):
            if ln.strip() == target_marker:
                offset = i
                break
        scan_lines = lines[offset:]
        scan_offset = offset
    else:
        scan_lines = lines
        scan_offset = 0

    indices: Dict[str, int] = {}
    for i, line in enumerate(scan_lines):
        for key, patterns in HEADING_PATTERNS.items():
            if key not in indices and line_matches_heading(line, patterns):
                indices[key] = i + scan_offset

    ordered_keys = ["1a", "1b", "2a", "2b", "sonuc"]
    sections: Dict[str, str] = {}
    for idx, key in enumerate(ordered_keys):
        start = indices.get(key)
        if start is None:
            continue
        later = [indices[k] for k in ordered_keys[idx + 1:] if k in indices and indices[k] > start]
        end = min(later) if later else len(lines)
        sections[key] = "\n".join(lines[start:end]).strip()
    return sections


def split_code_and_discussion(section_text: str) -> Tuple[str, str]:
    if not section_text:
        return "", ""
    lines = section_text.splitlines()
    code_start = None
    discuss_start = None
    # NOT: Marker tespiti satirin basinda olmak zorunda degil. Ogrenciler bazen
    # baslik formatini "1A. Convolution - MATLAB Code" / "1A. Discussion" gibi
    # yaziyor; "MATLAB Code" / "Discussion" kelimesi satirin sonunda kaliyor.
    # Bu yuzden re.search ile satirin herhangi bir yerinde ariyoruz, ayrica
    # uzun (icerik) cumleleriyle karismamasi icin kisa satir filtresi koyduk.
    code_marker_re = re.compile(r"\bmatlab\s+(?:kod\w*|code\w*)\b")
    discuss_marker_re = re.compile(r"\b(?:tartisma\w*|discussion\w*)\b")
    for i, line in enumerate(lines):
        nline = normalize_text(line).strip()
        # Baslik satirlari genelde kisa olur; uzun cumleler icindeki tesadufi
        # kelime gecislerini eslemiyoruz (false positive korumasi).
        if len(nline) > 80:
            continue
        if code_start is None and code_marker_re.search(nline):
            code_start = i + 1
        if discuss_start is None and discuss_marker_re.search(nline):
            discuss_start = i + 1

    # Sadece yorum metni olan bolumler (ozellikle sonuc)
    if code_start is None and discuss_start is None:
        return "", section_text.strip()

    if discuss_start is None:
        discuss_start = len(lines)

    if code_start is None:
        code_text = ""
    else:
        code_text = "\n".join(lines[code_start:discuss_start - 1 if discuss_start > 0 else discuss_start]).strip()

    discussion_text = "\n".join(lines[discuss_start:]).strip() if discuss_start < len(lines) else ""
    return code_text, discussion_text


def split_sentences(text: str) -> List[str]:
    # Join line-wrapped text (single newlines), preserve paragraph breaks (double)
    joined = re.sub(r"(?<!\n)\n(?!\n)", " ", text.replace("\r", ""))
    # Protect decimal numbers from being split
    protected = re.sub(r"(\d)\.(\d)", r"\1<DOT>\2", joined)
    raw = re.findall(r"[^.!?]+(?:[.!?]+|$)", protected)
    out = []
    for chunk in raw:
        c = re.sub(r"\s+", " ", chunk).strip().replace("<DOT>", ".")
        if c:
            out.append(c)
    return out


def count_sentences(text: str) -> int:
    return len([s for s in split_sentences(text) if len(s.split()) >= 4])


def extract_numbers(text: str) -> List[float]:
    vals = []
    for m in re.findall(r"(?<!\w)(?:\d+\.\d+|\d+)(?!\w)", text.replace(",", ".")):
        try:
            vals.append(float(m))
        except Exception:
            pass
    return vals


def sentences_with_keywords(text: str, keywords: List[str]) -> List[str]:
    out = []
    for s in split_sentences(text):
        ns = normalize_text(s)
        if any(k in ns for k in keywords):
            out.append(s)
    return out


def keyword_coverage(text: str, keyword_groups: List[List[str]]) -> int:
    ns = normalize_text(text)
    hits = 0
    for group in keyword_groups:
        if any(term in ns for term in group):
            hits += 1
    return hits


def choose_closest(numbers: List[float], expected: float) -> Optional[float]:
    if not numbers:
        return None
    return min(numbers, key=lambda x: abs(x - expected))


def approx_match(value: Optional[float], expected: float, *, abs_tol: float, rel_tol: float = 0.0) -> bool:
    if value is None:
        return False
    tol = max(abs_tol, abs(expected) * rel_tol)
    return abs(value - expected) <= tol


def safe_ratio(a: float, b: float) -> float:
    return a / b if b else 0.0


def plot_metrics(code_text: str) -> Dict[str, int]:
    tl = normalize_text(code_text)
    return {
        "plot": len(re.findall(r"\bplot\s*\(", tl)),
        "stem": len(re.findall(r"\bstem\s*\(", tl)),
        "xlabel": len(re.findall(r"xlabel\s*\(", tl)),
        "ylabel": len(re.findall(r"ylabel\s*\(", tl)),
        "title": len(re.findall(r"title\s*\(", tl)),
        "grid": len(re.findall(r"grid\s+on", tl)),
        "legend": len(re.findall(r"legend\s*\(", tl)),
        "hold": len(re.findall(r"hold\s+on", tl)),
        "marker": int(bool(re.search(r"marker|text\s*\(|annotation\s*\(|'o'|'\*'|scatter\s*\(", tl))),
        "figure": len(re.findall(r"\bfigure\b", tl)),
    }


# ---------------------------------------------------------------------------
# PDF GORSEL KONTROL
# ---------------------------------------------------------------------------

def image_hash_from_gray(gray: np.ndarray, size: int = 32) -> str:
    small = cv2.resize(gray, (size, size), interpolation=cv2.INTER_AREA)
    med = float(np.median(small))
    bits = (small > med).astype(np.uint8).flatten()
    return "".join("1" if b else "0" for b in bits)


def hamming_distance(a: str, b: str) -> int:
    if len(a) != len(b):
        return max(len(a), len(b))
    return sum(ch1 != ch2 for ch1, ch2 in zip(a, b))


def extract_figure_markers(page_text: str) -> List[int]:
    ns = normalize_text(page_text)
    figs: List[int] = []
    for n in range(1, 9):
        if re.search(rf"(?:grafik|figure|graph)\s*{n}\b", ns):
            figs.append(n)
    return figs


def analyze_pdf_visual(pdf_path: str) -> Dict[str, Any]:
    info: Dict[str, Any] = {
        "page_count": 0,
        "graph_like_pages": [],
        "screenshot_like_pages": [],
        "repeated_graph_pairs": [],
        "figure_pages": {},
        "figure_visual": {},
        "page_details": [],
    }
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        info["error"] = str(e)
        return info

    info["page_count"] = doc.page_count
    graph_hashes: List[Tuple[int, str]] = []

    for idx in range(doc.page_count):
        page = doc.load_page(idx)
        page_no = idx + 1
        page_text_raw = page.get_text("text")
        page_text = page_text_raw if isinstance(page_text_raw, str) else ""
        figures = extract_figure_markers(page_text)

        pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False)
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4:
            arr = cv2.cvtColor(arr, cv2.COLOR_RGBA2RGB)
        gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)

        edges = cv2.Canny(gray, 80, 180)
        edge_ratio = float(np.count_nonzero(edges)) / float(edges.size)

        h, w = gray.shape
        _, binimg = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        hk = np.ones((1, max(20, w // 10)), np.uint8)
        vk = np.ones((max(20, h // 10), 1), np.uint8)
        hor = cv2.morphologyEx(binimg, cv2.MORPH_OPEN, hk)
        ver = cv2.morphologyEx(binimg, cv2.MORPH_OPEN, vk)
        hline_ratio = float(np.count_nonzero(hor)) / float(h * w)
        vline_ratio = float(np.count_nonzero(ver)) / float(h * w)

        drawings = len(page.get_drawings())
        images = len(page.get_images(full=True))
        text_len = len(page_text.strip())

        # Sablonun talimat sayfalari (1-9) tablo cizgileri yuzunden "looks_graph"
        # olarak yanlis sinifllaniyor. Bu sayfalardaki yapilari karsilastirma
        # disinda tut. Talimat sayfa basliklari sayfa metninde gorunur.
        page_text_norm = normalize_text(page_text)
        is_instruction_page = any(m in page_text_norm for m in [
            "project objective and scope",
            "submission format and mandatory rules",
            "parameter determination",
            "mandatory report structure",
            "common student mistakes",
            "quick reference checklist",
            "how to write the conclusion section",
        ])

        graph_score = 0.0
        if images > 0:
            graph_score += 1.1
        if drawings >= 6:
            graph_score += 0.9
        elif drawings >= 2:
            graph_score += 0.5
        if edge_ratio >= 0.03:
            graph_score += 0.9
        elif edge_ratio >= 0.015:
            graph_score += 0.4
        if hline_ratio >= 0.002 or vline_ratio >= 0.0015:
            graph_score += 0.7
        if 40 <= text_len <= 2500:
            graph_score += 0.3
        if figures:
            graph_score += 0.5

        looks_graph = graph_score >= 1.7 and not is_instruction_page
        screenshot_like = images > 0 and drawings == 0 and edge_ratio >= 0.02 and not is_instruction_page

        record = {
            "page": page_no,
            "text_len": text_len,
            "images": images,
            "drawings": drawings,
            "edge_ratio": round(edge_ratio, 4),
            "hline_ratio": round(hline_ratio, 4),
            "vline_ratio": round(vline_ratio, 4),
            "graph_score": round(graph_score, 3),
            "looks_graph": looks_graph,
            "screenshot_like": screenshot_like,
            "figures": figures,
            "is_instruction_page": is_instruction_page,
        }
        info["page_details"].append(record)

        if looks_graph:
            info["graph_like_pages"].append(page_no)
            graph_hashes.append((page_no, image_hash_from_gray(gray)))
        if screenshot_like:
            info["screenshot_like_pages"].append(page_no)

        for fig in figures:
            info["figure_pages"].setdefault(fig, []).append(page_no)
            if looks_graph:
                info["figure_visual"][fig] = True

    page_meta = {d["page"]: d for d in info["page_details"]}
    for i in range(len(graph_hashes)):
        p1, h1 = graph_hashes[i]
        for j in range(i + 1, len(graph_hashes)):
            p2, h2 = graph_hashes[j]
            m1 = page_meta.get(p1, {})
            m2 = page_meta.get(p2, {})
            if max(m1.get("text_len", 0), m2.get("text_len", 0)) > 700:
                continue
            # Skip if both pages have figure markers with different numbers
            # (different MATLAB graphs naturally look similar)
            figs1 = set(m1.get("figures", []))
            figs2 = set(m2.get("figures", []))
            if figs1 and figs2 and not figs1.intersection(figs2):
                continue
            dist = hamming_distance(h1, h2)
            if dist <= 10:
                info["repeated_graph_pairs"].append((p1, p2, dist))

    doc.close()
    return info


def apply_visual_evidence(criteria: Dict[str, "CriterionResult"], visual_info: Dict[str, Any], notes: Dict[str, str]) -> None:
    fig_to_key = {
        1: "1A_graf_xt", 2: "1A_graf_ht", 3: "1A_graf_yt", 4: "1B_graf_ucspek",
        5: "2A_graf_xn", 6: "2A_graf_hn", 7: "2A_graf_kars", 8: "2B_graf_spek",
    }
    figure_pages = visual_info.get("figure_pages", {}) or {}
    figure_visual = visual_info.get("figure_visual", {}) or {}

    for fig, key in fig_to_key.items():
        if key not in criteria:
            continue
        cur = criteria[key]
        pages = figure_pages.get(fig, [])
        has_marker = bool(pages)
        has_visual = bool(figure_visual.get(fig, False))

        if has_visual and cur.score < cur.max_score:
            boosted = min(cur.max_score, cur.score + 1)
            criteria[key] = make_result(
                boosted, cur.max_score, min(1.0, cur.confidence + 0.12),
                cur.reason + f"; gorsel_kanit=Grafik {fig} sayfa {pages}"
            )
        elif has_marker and cur.score == 0:
            criteria[key] = make_result(
                1, cur.max_score, min(0.6, cur.confidence + 0.08),
                cur.reason + f"; baslik_kaniti=Grafik {fig} sayfa {pages}"
            )

    repeated_pairs = visual_info.get("repeated_graph_pairs", []) or []
    if repeated_pairs:
        pairs_txt = ", ".join([f"{a}-{b}(d={d})" for a, b, d in repeated_pairs[:5]])
        notes["GRAFIK_TEKRAR?"] = f"Benzer grafik sayfalari: {pairs_txt}"

# ---------------------------------------------------------------------------
# KRITER SONUC MODELI
# ---------------------------------------------------------------------------

@dataclass
class CriterionResult:
    score: int
    max_score: int
    confidence: float
    reason: str


def clamp_score(score: float, max_score: int) -> int:
    return max(0, min(max_score, int(round(score))))


def make_result(raw_score: float, max_score: int, confidence: float, reason: str) -> CriterionResult:
    return CriterionResult(
        score=clamp_score(raw_score, max_score),
        max_score=max_score,
        confidence=max(0.0, min(1.0, confidence)),
        reason=reason.strip(),
    )


# ---------------------------------------------------------------------------
# AKILLI DEGERLENDIRME ALT FONKSIYONLARI
# ---------------------------------------------------------------------------

def evaluate_1a(code_text: str, discussion_text: str, exp: Dict[str, float]) -> Dict[str, CriterionResult]:
    code = normalize_text(code_text)
    disc = normalize_text(discussion_text)
    pm = plot_metrics(code_text)
    labels = min(pm["xlabel"], pm["ylabel"], pm["title"], pm["grid"])
    results: Dict[str, CriterionResult] = {}

    has_sin = "sin(2*pi" in code.replace(" ", "") or "sin(2*pi*f0*t" in code.replace(" ", "")
    has_exp = "exp(-a" in code or "exp(- a" in code
    score = 3 if (has_sin and has_exp) else (1 if (has_sin or has_exp) else 0)
    conf = 0.95 if score == 3 else 0.55 if score > 0 else 0.2
    results["1A_kod_sinyal"] = make_result(score, 3, conf, f"sin={has_sin}, exp={has_exp}")

    has_conv = "conv(" in code
    has_dt_mult = "*dt" in code.replace(" ", "") or bool(re.search(r"conv\s*\([^\n]+\)\s*\*\s*dt", code))
    score = 3 if (has_conv and has_dt_mult) else (2 if has_conv else 0)
    conf = 0.95 if score == 3 else 0.65 if score == 2 else 0.2
    results["1A_kod_conv"] = make_result(score, 3, conf, f"conv={has_conv}, dt_carpimi={has_dt_mult}")

    has_time_axis = ("length(" in code and ("dt" in code or "t_y" in code or "ty" in code)) or ("0:dt" in code.replace(" ", ""))
    score = 2 if has_time_axis else 0
    conf = 0.85 if has_time_axis else 0.25
    results["1A_kod_texsen"] = make_result(score, 2, conf, f"zaman_ekseni={has_time_axis}")

    score_xt = 2 if (pm["plot"] >= 1 and labels >= 1) else (1 if pm["plot"] >= 1 else 0)
    score_ht = 2 if (pm["plot"] >= 2 and labels >= 2) else (1 if pm["plot"] >= 2 else 0)
    score_yt = 3 if (pm["plot"] >= 3 and labels >= 3 and pm["marker"]) else (2 if pm["plot"] >= 3 else 0)
    results["1A_graf_xt"] = make_result(score_xt, 2, 0.85 if score_xt == 2 else 0.55 if score_xt else 0.2, f"plot={pm['plot']}, etiket={labels}")
    results["1A_graf_ht"] = make_result(score_ht, 2, 0.85 if score_ht == 2 else 0.55 if score_ht else 0.2, f"plot={pm['plot']}, etiket={labels}")
    results["1A_graf_yt"] = make_result(score_yt, 3, 0.9 if score_yt >= 2 else 0.25, f"plot={pm['plot']}, marker={bool(pm['marker'])}, etiket={labels}")

    nums = extract_numbers(discussion_text)
    peak_kw = any(k in disc for k in ["tepe", "peak", "maksimum", "maximum"])
    amp_val = choose_closest(nums, exp["y_peak"])
    time_val = choose_closest(nums, exp["y_peak_t"])
    amp_ok = approx_match(amp_val, exp["y_peak"], abs_tol=0.03, rel_tol=0.12)
    time_ok = approx_match(time_val, exp["y_peak_t"], abs_tol=0.05, rel_tol=0.2)
    raw = 0
    if peak_kw:
        raw += 2
    if len(nums) >= 2:
        raw += 1
    if amp_ok:
        raw += 1
    if time_ok:
        raw += 1
    conf = 0.35 + 0.25 * int(peak_kw) + 0.2 * int(amp_ok) + 0.2 * int(time_ok)
    results["1A_tart_tepe"] = make_result(raw, 5, conf, f"peak_kw={peak_kw}, tepe~{amp_val}, zaman~{time_val}, beklenen=({exp['y_peak']}, {exp['y_peak_t']})")

    coverage = keyword_coverage(discussion_text, [
        ["uzun", "uzar", "uzama", "longer", "duration", "sure", "extends", "outlasts"],
        ["sonsuz", "hafiza", "zaman sabiti", "time constant", "filtre hafizasi", "birikim", "memory", "accumulation", "energy storage", "infinite impulse"],
    ])
    raw = 2.5 * coverage
    conf = 0.3 + 0.3 * coverage
    results["1A_tart_uzama"] = make_result(raw, 5, conf, f"kapsam={coverage}/2")

    has_a_scenario = any(k in disc for k in ["2 kat", "iki kat", "double", "a degeri", "a art", "a iki", "doubled", "twice", "increasing a", "if a were"])
    shorter = any(k in disc for k in ["kisa", "kisal", "daha kisa", "shorter", "daha hizli", "faster decay", "decays faster", "shorter duration"]) 
    lower_peak = any(k in disc for k in ["tepe azal", "peak azal", "peak dus", "daha kucuk tepe", "genlik azal", "attenuation art", "lower peak", "peak decreases", "amplitude decreases", "reduced peak", "smaller peak"]) 
    raw = 0
    if has_a_scenario:
        raw += 2
    if shorter:
        raw += 1.5
    if lower_peak:
        raw += 1.5
    if keyword_coverage(discussion_text, [["fiziksel", "cunku", "because", "filtre", "sonum", "enerji birikimi", "physically", "damping", "energy accumulation"]]) > 0:
        raw += 0.5
    conf = 0.3 + 0.25 * int(has_a_scenario) + 0.2 * int(shorter) + 0.2 * int(lower_peak)
    results["1A_tart_a_deg"] = make_result(raw, 5, conf, f"senaryo={has_a_scenario}, sure_kisalir={shorter}, tepe_azalir={lower_peak}")

    return results


def evaluate_1b(code_text: str, discussion_text: str, exp: Dict[str, float]) -> Dict[str, CriterionResult]:
    code = normalize_text(code_text)
    disc = normalize_text(discussion_text)
    pm = plot_metrics(code_text)
    results: Dict[str, CriterionResult] = {}

    has_fft = bool(re.search(r"f\s*f\s*t\s*\(", code))
    has_abs = bool(re.search(r"abs\s*\(\s*f\s*f\s*t", code.replace(" ", ""))) or bool(re.search(r"abs\s*\(\s*fft", code))
    has_norm = bool(re.search(r"/\s*(length|nfft|n|numel)", code))
    score = 3 if (has_fft and (has_abs or has_norm)) else (1 if has_fft else 0)
    results["1B_kod_fft"] = make_result(score, 3, 0.95 if score == 3 else 0.5 if score else 0.2, f"fft={has_fft}, abs/norm={has_abs or has_norm}")

    has_faxis = bool(re.search(r"\(0:.*\)\*fs/|fs/n|1/dt|1000", code))
    results["1B_kod_faxis"] = make_result(2 if has_faxis else 0, 2, 0.85 if has_faxis else 0.25, f"faxis={has_faxis}")

    has_half = bool(re.search(r"fs/2|half|n/2|floor\(.*?/2\)", code))
    results["1B_kod_tektar"] = make_result(2 if has_half else 0, 2, 0.8 if has_half else 0.25, f"tek_taraf={has_half}")

    multi_plot = pm["plot"] >= 3 or (pm["hold"] >= 1 and pm["legend"] >= 1)
    score = 5 if (multi_plot and pm["legend"] >= 1) else (3 if pm["legend"] >= 1 else (1 if has_fft else 0))
    conf = 0.9 if score == 5 else 0.65 if score >= 3 else 0.3
    results["1B_graf_ucspek"] = make_result(score, 5, conf, f"plot={pm['plot']}, legend={pm['legend']}, hold={pm['hold']}")

    nums = extract_numbers(discussion_text)
    dom_kw = any(k in disc for k in ["baskin", "dominant", "f0", "temel frekans", "fundamental"])
    dom_val = choose_closest(nums, exp["dominant_f"])
    dom_ok = approx_match(dom_val, exp["dominant_f"], abs_tol=0.25, rel_tol=0.05)
    raw = 0
    if dom_kw:
        raw += 2
    if dom_ok:
        raw += 2
    elif len(nums) > 0:
        raw += 1
    results["1B_tart_baskin"] = make_result(raw, 4, 0.35 + 0.3 * int(dom_kw) + 0.35 * int(dom_ok), f"dominant~{dom_val}, beklenen={exp['dominant_f']}")

    low_pass = any(k in disc for k in ["alcak geciren", "low pass", "low-pass", "lpf", "lowpass"])
    reason = keyword_coverage(discussion_text, [["dusuk frekans", "low freq", "0 hz", "dc"], ["yuksek frekans", "high freq", "zayif", "azal", "attenuat", "suppress"]])
    raw = 2 if low_pass else 0
    raw += min(2, reason)
    results["1B_tart_filtre"] = make_result(raw, 4, 0.35 + 0.35 * int(low_pass) + 0.15 * reason, f"low_pass={low_pass}, aciklama_kapsami={reason}")

    comparison_cov = keyword_coverage(discussion_text, [
        ["y(f)", "yf", "cikis", "output"],
        ["x(f)", "xf", "giris", "input"],
        ["zayif", "attenuat", "bastir", "suppress"],
        ["korun", "remain", "temel bilesen", "preserv", "dominant freq", "fundamental"],
    ])
    raw = min(5, 1.25 * comparison_cov)
    results["1B_tart_yxkars"] = make_result(raw, 5, 0.3 + 0.15 * comparison_cov, f"kapsam={comparison_cov}/4")
    return results


def evaluate_2a(code_text: str, discussion_text: str, exp: Dict[str, float], f0: int) -> Dict[str, CriterionResult]:
    code = normalize_text(code_text)
    disc = normalize_text(discussion_text)
    pm = plot_metrics(code_text)
    labels = min(pm["xlabel"], pm["ylabel"], pm["title"], pm["grid"])
    results: Dict[str, CriterionResult] = {}

    has_ts = bool(re.search(r"ts\s*=\s*0\.01|0\.01", code))
    has_discrete_x = ("x_n" in code) or ("sin(2*pi*f0*n" in code.replace(" ", ""))
    score = 3 if (has_ts and has_discrete_x) else (1 if has_ts else 0)
    results["2A_kod_ornekle"] = make_result(score, 3, 0.9 if score == 3 else 0.55 if score else 0.2, f"Ts={has_ts}, x[n]={has_discrete_x}")

    has_conv = "conv(" in code
    wrong_dt = bool(re.search(r"conv\s*\([^)]*\)\s*\*\s*dt", code.replace(" ", "")))
    score = 2 if (has_conv and not wrong_dt) else (1 if has_conv else 0)
    results["2A_kod_ayrconv"] = make_result(score, 2, 0.85 if score == 2 else 0.45 if score else 0.2, f"conv={has_conv}, hatali_dt={wrong_dt}")

    score_xn = 2 if (pm["stem"] >= 1 and labels >= 1) else (1 if pm["stem"] >= 1 else 0)
    score_hn = 2 if (pm["stem"] >= 2 and labels >= 2) else (1 if pm["stem"] >= 2 else 0)
    compare_ok = (pm["plot"] >= 1 and pm["stem"] >= 1 and pm["hold"] >= 1 and pm["legend"] >= 1)
    score_k = 3 if compare_ok else (1 if pm["hold"] >= 1 else 0)
    results["2A_graf_xn"] = make_result(score_xn, 2, 0.85 if score_xn == 2 else 0.5 if score_xn else 0.2, f"stem={pm['stem']}, etiket={labels}")
    results["2A_graf_hn"] = make_result(score_hn, 2, 0.85 if score_hn == 2 else 0.5 if score_hn else 0.2, f"stem={pm['stem']}, etiket={labels}")
    results["2A_graf_kars"] = make_result(score_k, 3, 0.9 if score_k == 3 else 0.45 if score_k else 0.2, f"plot={pm['plot']}, stem={pm['stem']}, hold={pm['hold']}, legend={pm['legend']}")

    scale_cov = keyword_coverage(discussion_text, [["genlik", "amplitude", "olcek", "scale", "magnitude"], ["ts", "0.01", "carp", "scaling", "multiply"]])
    results["2A_tart_genlik"] = make_result(1.5 * scale_cov, 3, 0.3 + 0.25 * scale_cov, f"kapsam={scale_cov}/2")

    nyq_sent = " ".join(sentences_with_keywords(discussion_text, ["nyquist", "yeterli", "sufficient", "fs", "100", "2*f0", "2f0"]))
    nyq_nums = extract_numbers(nyq_sent)
    has_100 = any(abs(x - 100) < 1e-6 for x in nyq_nums)
    has_2f0 = any(abs(x - exp["two_f0"]) <= 0.5 for x in nyq_nums)
    says_ok = any(k in normalize_text(nyq_sent) for k in [
        "yeterli", "suffic", "sufci", "saglar", "uygun",
        "aliasing olmaz", "problem yok", "satisfied", "met",
        "no aliasing", "without aliasing",
    ])
    raw = 0
    if says_ok == exp["nyquist_ok"]:
        raw += 1.5
    if has_100:
        raw += 0.75
    if has_2f0 or f"{f0}" in normalize_text(nyq_sent):
        raw += 0.75
    results["2A_tart_nyquist"] = make_result(raw, 3, 0.35 + 0.25 * int(says_ok == exp['nyquist_ok']) + 0.2 * int(has_100) + 0.2 * int(has_2f0), f"nyquist_yorumu={says_ok}, 100_var={has_100}, 2f0_var={has_2f0}, beklenen={exp['nyquist_ok']}")

    alias_sent = " ".join(sentences_with_keywords(discussion_text, ["alias", "ts = 0.1", "0.1", "fs = 10", "10 hz", "bozulma", "distortion"]))
    alias_ns = normalize_text(alias_sent)
    says_alias = any(k in alias_ns for k in ["alias", "olur", "bozulma olur", "distortion olur", "undersample", "would occur", "violated", "insufficient"])
    says_no_alias = any(k in alias_ns for k in ["olmaz", "yok", "no alias", "aliasing olmaz"])
    expected_alias = exp["alias_if_ts_0_1"]
    verdict_ok = (expected_alias and says_alias) or ((not expected_alias) and says_no_alias)
    raw = 0
    if verdict_ok:
        raw += 1.5
    if "10" in alias_ns or "0.1" in alias_ns:
        raw += 0.5
    results["2A_tart_alias"] = make_result(raw, 2, 0.35 + 0.35 * int(verdict_ok) + 0.1 * int("10" in alias_ns), f"alias_yorumu={alias_sent[:120] or '-'}, beklenen_alias={expected_alias}")
    return results


def evaluate_2b(code_text: str, discussion_text: str, exp: Dict[str, float]) -> Dict[str, CriterionResult]:
    code = normalize_text(code_text)
    disc = normalize_text(discussion_text)
    pm = plot_metrics(code_text)
    results: Dict[str, CriterionResult] = {}

    has_fft = bool(re.search(r"f\s*f\s*t\s*\(", code))
    results["2B_kod_fft"] = make_result(2 if has_fft else 0, 2, 0.85 if has_fft else 0.2, f"fft={has_fft}")

    has_faxis = bool(re.search(r"fs_d|1/ts|100|\(0:.*\)\*fs_d/", code))
    results["2B_kod_faxis"] = make_result(2 if has_faxis else 0, 2, 0.8 if has_faxis else 0.2, f"faxis={has_faxis}")

    score = 3 if (pm["legend"] >= 1 and (pm["plot"] >= 2 or pm["hold"] >= 1)) else (1 if has_fft else 0)
    results["2B_graf_spek"] = make_result(score, 3, 0.85 if score == 3 else 0.45 if score else 0.2, f"plot={pm['plot']}, legend={pm['legend']}, hold={pm['hold']}")

    nums = extract_numbers(discussion_text)
    dom_val = choose_closest(nums, exp["dominant_f"])
    dom_ok = approx_match(dom_val, exp["dominant_f"], abs_tol=0.25, rel_tol=0.05)
    similar_kw = any(k in disc for k in ["benzer", "similar", "ayni", "same", "uyumlu", "consistent", "matches", "agrees"])
    raw = 0
    if similar_kw:
        raw += 1.5
    if dom_ok:
        raw += 1.5
    results["2B_tart_benzer"] = make_result(raw, 3, 0.35 + 0.25 * int(similar_kw) + 0.35 * int(dom_ok), f"benzerlik={similar_kw}, dominant~{dom_val}, beklenen={exp['dominant_f']}")

    res_sent = " ".join(sentences_with_keywords(discussion_text, ["cozunur", "resolution", "df", "delta f", "frequency resolution"]))
    res_nums = extract_numbers(res_sent)
    cand_cont = choose_closest(res_nums, exp["df_cont"])
    cand_disc = choose_closest(res_nums, exp["df_disc"])
    num_ok = approx_match(cand_cont, exp["df_cont"], abs_tol=0.1, rel_tol=0.25) or approx_match(cand_disc, exp["df_disc"], abs_tol=0.1, rel_tol=0.25)
    raw = 1 if len(res_sent) > 0 else 0
    if num_ok:
        raw += 1
    results["2B_tart_cozunur"] = make_result(raw, 2, 0.35 + 0.25 * int(len(res_sent) > 0) + 0.3 * int(num_ok), f"df_adayi={res_nums[:4]}, beklenen=({exp['df_cont']}, {exp['df_disc']})")

    filt_cov = keyword_coverage(discussion_text, [["alcak geciren", "low pass", "filter", "filtre"], ["zayif", "attenuat", "bastir", "suppress"], ["surekli ile tutarli", "consistent", "uyumlu", "confirms", "agrees"]])
    results["2B_tart_filtre"] = make_result(filt_cov, 3, 0.3 + 0.2 * filt_cov, f"kapsam={filt_cov}/3")
    return results


def evaluate_summary(discussion_text: str, exp: Dict[str, float]) -> Dict[str, CriterionResult]:
    disc = normalize_text(discussion_text)
    nums = extract_numbers(discussion_text)
    results: Dict[str, CriterionResult] = {}

    peak_cont = choose_closest(nums, exp["y_peak"])
    peak_disc = choose_closest(nums, exp["y_d_peak_scaled"])
    dur_cont = choose_closest(nums, exp["y_duration"])
    dur_disc = choose_closest(nums, exp["y_d_duration"])
    good_nums = sum([
        approx_match(peak_cont, exp["y_peak"], abs_tol=0.03, rel_tol=0.12),
        approx_match(peak_disc, exp["y_d_peak_scaled"], abs_tol=0.05, rel_tol=0.18),
        approx_match(dur_cont, exp["y_duration"], abs_tol=0.15, rel_tol=0.15),
        approx_match(dur_disc, exp["y_d_duration"], abs_tol=0.15, rel_tol=0.15),
    ])
    conv_cov = keyword_coverage(discussion_text, [["tepe", "peak", "maximum"], ["sure", "duration", "length"], ["uyum", "match", "benzer", "consistent", "agreement"]])
    raw = min(3, 0.75 * good_nums + 0.5 * conv_cov)
    results["SN_konv_kars"] = make_result(raw, 3, 0.3 + 0.1 * good_nums + 0.15 * conv_cov, f"sayisal_eslesme={good_nums}/4, kapsam={conv_cov}/3")

    fft_cov = keyword_coverage(discussion_text, [["cozunur", "resolution", "df", "frequency resolution"], ["maksimum frekans", "nyquist", "fs/2", "maximum frequency"], ["surekli", "ayrik", "fft", "continuous", "discrete"]])
    results["SN_fft_kars"] = make_result(min(3, fft_cov), 3, 0.3 + 0.18 * fft_cov, f"kapsam={fft_cov}/3")

    atten_val = choose_closest(nums, exp["attenuation_pct"])
    atten_ok = approx_match(atten_val, exp["attenuation_pct"], abs_tol=3.0, rel_tol=0.08)
    hf0_cov = keyword_coverage(discussion_text, [["h(f0)", "hf0", "|h(f0)|"], ["yuzde", "percent", "%", "percentage"], ["zayif", "attenuat", "azal", "suppress", "reduction"]])
    raw = min(4, hf0_cov + 1 * int(atten_ok))
    results["SN_hf0_hesap"] = make_result(raw, 4, 0.3 + 0.15 * hf0_cov + 0.2 * int(atten_ok), f"zayiflatma~{atten_val}, beklenen={exp['attenuation_pct']}")
    return results


def enforce_discussion_floor(criteria: Dict[str, CriterionResult], discussion_text: str, affected_keys: List[str]) -> None:
    sent_count = count_sentences(discussion_text)
    word_count = len(discussion_text.split())
    if sent_count < 2 and word_count < 40:
        for key in affected_keys:
            cur = criteria[key]
            criteria[key] = make_result(min(cur.score, 1), cur.max_score, min(cur.confidence, 0.45), f"Kisa tartisma: cumle~{sent_count}, kelime={word_count}. {cur.reason}")


# ---------------------------------------------------------------------------
# GENEL PDF DEGERLENDIRME
# ---------------------------------------------------------------------------

def evaluate_pdf(text: str, filename: str, expected_params: Dict[str, int], expected_results: Dict[str, float], visual_info: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    sections = find_section_boundaries(text)
    criteria: Dict[str, CriterionResult] = {}
    notes: Dict[str, str] = {}

    # Dosya adi formati kontrolu. LMS dosyalari sistemin verdigi adla geldigi
    # icin varsayilanda (PENALIZE_FILENAME=False) ceza uygulanmiyor; sadece
    # bilgi amacli not eklenir. PENALIZE_FILENAME=True yaparsan eski -5 ceza.
    fn_ok = bool(re.match(r"^\d{8,12}_[A-Za-zÇçĞğİıÖöŞşÜı][A-Za-zÇçĞğİıÖöŞşÜı_-]*_(?:Proje|Project)\.pdf$", filename, re.I))
    if PENALIZE_FILENAME:
        ceza_dosya_adi = 0 if fn_ok else -5
    else:
        ceza_dosya_adi = 0
        if not fn_ok:
            notes["DOSYA_ADI"] = f"Dosya adi standart formatta degil: {filename} (LMS adlandirmasi olabilir)"

    # Sablon temizligi cezasi: ogrenci 1-9. sayfa talimatlarini PDF'e dahil
    # etmisse -5 ceza (dosya adi kurali ile ayni kategoride: gonderim formati).
    template_uncleaned, template_markers = detect_template_uncleaned(text)
    ceza_sablon = -5 if template_uncleaned else 0
    if template_uncleaned:
        notes["SABLON?"] = f"Sablon talimat sayfalari silinmemis (dahil edilen marker sayisi: {len(template_markers)})"

    reported = extract_reported_params(text)
    param_ok = all(reported.get(k) == expected_params[k] for k in ["a", "f0", "T"])
    if not param_ok:
        notes["PARAM"] = f"Bulunan={reported}, Beklenen={expected_params}"

    # 1A
    sec = sections.get("1a", "")
    if sec:
        code, disc = split_code_and_discussion(sec)
        block = evaluate_1a(code, disc, expected_results)
        enforce_discussion_floor(block, disc, ["1A_tart_tepe", "1A_tart_uzama", "1A_tart_a_deg"])
        criteria.update(block)
    else:
        notes["B1A"] = "Bolum bulunamadi"
        for key, max_score in [("1A_kod_sinyal",3),("1A_kod_conv",3),("1A_kod_texsen",2),("1A_graf_xt",2),("1A_graf_ht",2),("1A_graf_yt",3),("1A_tart_tepe",5),("1A_tart_uzama",5),("1A_tart_a_deg",5)]:
            criteria[key] = make_result(0, max_score, 0.0, "Bolum bulunamadi")

    # 1B
    sec = sections.get("1b", "")
    if sec:
        code, disc = split_code_and_discussion(sec)
        block = evaluate_1b(code, disc, expected_results)
        enforce_discussion_floor(block, disc, ["1B_tart_baskin", "1B_tart_filtre", "1B_tart_yxkars"])
        criteria.update(block)
    else:
        notes["B1B"] = "Bolum bulunamadi"
        for key, max_score in [("1B_kod_fft",3),("1B_kod_faxis",2),("1B_kod_tektar",2),("1B_graf_ucspek",5),("1B_tart_baskin",4),("1B_tart_filtre",4),("1B_tart_yxkars",5)]:
            criteria[key] = make_result(0, max_score, 0.0, "Bolum bulunamadi")

    # 2A
    sec = sections.get("2a", "")
    if sec:
        code, disc = split_code_and_discussion(sec)
        block = evaluate_2a(code, disc, expected_results, expected_params["f0"])
        enforce_discussion_floor(block, disc, ["2A_tart_genlik", "2A_tart_nyquist", "2A_tart_alias"])
        criteria.update(block)
    else:
        notes["B2A"] = "Bolum bulunamadi"
        for key, max_score in [("2A_kod_ornekle",3),("2A_kod_ayrconv",2),("2A_graf_xn",2),("2A_graf_hn",2),("2A_graf_kars",3),("2A_tart_genlik",3),("2A_tart_nyquist",3),("2A_tart_alias",2)]:
            criteria[key] = make_result(0, max_score, 0.0, "Bolum bulunamadi")

    # 2B
    sec = sections.get("2b", "")
    if sec:
        code, disc = split_code_and_discussion(sec)
        block = evaluate_2b(code, disc, expected_results)
        enforce_discussion_floor(block, disc, ["2B_tart_benzer", "2B_tart_cozunur", "2B_tart_filtre"])
        criteria.update(block)
    else:
        notes["B2B"] = "Bolum bulunamadi"
        for key, max_score in [("2B_kod_fft",2),("2B_kod_faxis",2),("2B_graf_spek",3),("2B_tart_benzer",3),("2B_tart_cozunur",2),("2B_tart_filtre",3)]:
            criteria[key] = make_result(0, max_score, 0.0, "Bolum bulunamadi")

    # Sonuc
    sec = sections.get("sonuc", "")
    if sec:
        _, disc = split_code_and_discussion(sec)
        block = evaluate_summary(disc, expected_results)
        enforce_discussion_floor(block, disc, ["SN_konv_kars", "SN_fft_kars", "SN_hf0_hesap"])
        criteria.update(block)
    else:
        notes["Sonuc"] = "Bolum bulunamadi"
        for key, max_score in [("SN_konv_kars",3),("SN_fft_kars",3),("SN_hf0_hesap",4)]:
            criteria[key] = make_result(0, max_score, 0.0, "Bolum bulunamadi")

    # Gorsel katman: grafik puanlarini gerekirse toparla
    visual_info = visual_info or {}
    apply_visual_evidence(criteria, visual_info, notes)

    # Grafik etiketi cezasi: daha yumuşak, sadece ciddi eksiklerde
    code_sections = []
    for key in ["1a", "1b", "2a", "2b"]:
        if key in sections:
            code, _ = split_code_and_discussion(sections[key])
            code_sections.append(plot_metrics(code))
    total_xlabel = sum(m["xlabel"] for m in code_sections)
    total_ylabel = sum(m["ylabel"] for m in code_sections)
    total_title = sum(m["title"] for m in code_sections)
    total_grid = sum(m["grid"] for m in code_sections)
    total_labeled = min(total_xlabel, total_ylabel, total_title, total_grid)
    approx_figures = max(total_title, total_xlabel, total_ylabel)
    missing = max(0, approx_figures - total_labeled)
    ceza_etiket = -1 if missing >= 3 else 0

    text_len = len(text)
    visual_graph_pages = len(visual_info.get("graph_like_pages", []) or [])
    screenshot_pages = len(visual_info.get("screenshot_like_pages", []) or [])
    repeated_pairs = len(visual_info.get("repeated_graph_pairs", []) or [])
    screenshot_uyari = (text_len < 1200) or ((text_len < 3500) and (screenshot_pages >= max(3, (visual_info.get("page_count", 0) // 2))))
    if screenshot_uyari:
        notes["SCREENSHOT?"] = f"Metin/gorsel dengesi supheli: karakter={text_len}, screenshot_benzeri_sayfa={screenshot_pages}"

    # Toplamlar
    b1a_total = min(sum(c.score for k, c in criteria.items() if k.startswith("1A_")), 30)
    b1b_total = min(sum(c.score for k, c in criteria.items() if k.startswith("1B_")), 25)
    b2a_total = min(sum(c.score for k, c in criteria.items() if k.startswith("2A_")), 20)
    b2b_total = min(sum(c.score for k, c in criteria.items() if k.startswith("2B_")), 15)
    sn_total = min(sum(c.score for k, c in criteria.items() if k.startswith("SN_")), 10)
    ham_toplam = b1a_total + b1b_total + b2a_total + b2b_total + sn_total
    cezalar = ceza_dosya_adi + ceza_etiket + ceza_sablon
    net_toplam = max(0, ham_toplam + cezalar)

    confidence = safe_ratio(sum(c.confidence for c in criteria.values()), max(1, len(criteria)))
    if visual_graph_pages >= 4:
        confidence = min(1.0, confidence + 0.03)
    manual_review = confidence < 0.6 or screenshot_uyari or not param_ok or repeated_pairs >= 4

    # v2'deki alanlara geriye uyumluluk icin puan map'i
    score_map = {k: c.score for k, c in criteria.items()}
    confidence_map = {k: round(c.confidence, 3) for k, c in criteria.items()}
    reason_map = {k: c.reason for k, c in criteria.items()}

    return {
        "criteria": criteria,
        "scores": score_map,
        "confidences": confidence_map,
        "reasons": reason_map,
        "notes": notes,
        "b1a": b1a_total,
        "b1b": b1b_total,
        "b2a": b2a_total,
        "b2b": b2b_total,
        "sonuc": sn_total,
        "ham_toplam": ham_toplam,
        "ceza_dosya_adi": ceza_dosya_adi,
        "ceza_etiket": ceza_etiket,
        "ceza_sablon": ceza_sablon,
        "template_uncleaned": template_uncleaned,
        "net_toplam": net_toplam,
        "param_ok": param_ok,
        "screenshot_uyari": screenshot_uyari,
        "overall_confidence": round(confidence, 3),
        "manual_review": manual_review,
        "reported_params": reported,
        "visual_info": visual_info,
        "graph_like_pages": visual_graph_pages,
        "screenshot_like_pages": screenshot_pages,
        "repeated_graph_pairs": repeated_pairs,
    }


# ---------------------------------------------------------------------------
# EXCEL RAPORLAMA
# ---------------------------------------------------------------------------

def create_excel(results: List[Dict[str, Any]], output_path: str) -> str:
    wb = Workbook()

    header_font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1B2A4A")
    sub_header_fill = PatternFill("solid", fgColor="2D5A9E")
    ok_fill = PatternFill("solid", fgColor="D4EDDA")
    warn_fill = PatternFill("solid", fgColor="FFF3CD")
    bad_fill = PatternFill("solid", fgColor="F8D7DA")
    light_gray = PatternFill("solid", fgColor="F5F5F5")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    bold_font = Font(name="Arial", bold=True, size=10)
    normal_font = Font(name="Arial", size=10)

    # Sayfa 1
    ws = wb.active
    assert ws is not None
    ws.title = "Notlar"
    headers = [
        ("Sira", 5), ("Ogrenci No", 15), ("Ad Soyad", 20), ("N", 5),
        ("a", 4), ("f0", 4), ("T", 4), ("Param?", 8),
        ("1A\n(30)", 8), ("1B\n(25)", 8), ("2A\n(20)", 8), ("2B\n(15)", 8), ("Sonuc\n(10)", 8),
        ("Ham", 8), ("Ceza", 7), ("Net", 8),
        ("Guven", 8), ("GrafikSf", 8), ("Tekrar?", 8), ("Manuel?", 9), ("Notlar", 38),
    ]
    for col_idx, (h, w) in enumerate(headers, 1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[1].height = 34

    for i, r in enumerate(results, 1):
        ev = r["eval"]
        notes = "; ".join([f"{k}: {v}" for k, v in ev["notes"].items()]) if ev["notes"] else ""
        row = [
            i, r["ogr_no"] or "?", r["ogr_name"] or "?", r["N"] if r["N"] is not None else "?",
            r["params"]["a"] if r["params"] else "?", r["params"]["f0"] if r["params"] else "?", r["params"]["T"] if r["params"] else "?",
            "OK" if ev["param_ok"] else "HATALI",
            ev["b1a"], ev["b1b"], ev["b2a"], ev["b2b"], ev["sonuc"],
            ev["ham_toplam"], ev["ceza_dosya_adi"] + ev["ceza_etiket"] + ev.get("ceza_sablon", 0), ev["net_toplam"],
            ev["overall_confidence"], ev.get("graph_like_pages", 0), ev.get("repeated_graph_pairs", 0), "EVET" if ev["manual_review"] else "-", notes,
        ]
        for col_idx, val in enumerate(row, 1):
            c = ws.cell(row=i + 1, column=col_idx, value=val)
            c.font = normal_font
            c.border = thin_border
            c.alignment = left_wrap if col_idx in (2, 3, 21) else center

        if i % 2 == 0:
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row=i + 1, column=col_idx).fill = light_gray

        ws.cell(row=i + 1, column=8).fill = ok_fill if ev["param_ok"] else bad_fill
        net_cell = ws.cell(row=i + 1, column=16)
        net_cell.font = Font(name="Arial", bold=True, size=11)
        if ev["net_toplam"] >= 70:
            net_cell.fill = ok_fill
        elif ev["net_toplam"] >= 50:
            net_cell.fill = warn_fill
        else:
            net_cell.fill = bad_fill

        conf_cell = ws.cell(row=i + 1, column=17)
        conf = ev["overall_confidence"]
        if conf >= 0.8:
            conf_cell.fill = ok_fill
        elif conf >= 0.6:
            conf_cell.fill = warn_fill
        else:
            conf_cell.fill = bad_fill

        rep_cell = ws.cell(row=i + 1, column=19)
        if ev.get("repeated_graph_pairs", 0):
            rep_cell.fill = warn_fill
        if ev["manual_review"]:
            ws.cell(row=i + 1, column=20).fill = warn_fill

    last_row = len(results) + 2
    ws.cell(row=last_row, column=2, value="ORTALAMA").font = bold_font
    for col_idx in [9, 10, 11, 12, 13, 14, 16, 17, 18, 19]:
        col_letter = get_column_letter(col_idx)
        c = ws.cell(row=last_row, column=col_idx)
        c.value = f"=AVERAGE({col_letter}2:{col_letter}{last_row-1})"
        c.font = bold_font
        c.alignment = center
        c.number_format = "0.0"
        c.border = thin_border
    ws.auto_filter.ref = f"A1:U{len(results)+1}"
    ws.freeze_panes = "A2"

    # Sayfa 2 - detay
    ws2 = wb.create_sheet("Detayli Puanlar")
    score_keys = [
        "1A_kod_sinyal", "1A_kod_conv", "1A_kod_texsen", "1A_graf_xt", "1A_graf_ht", "1A_graf_yt", "1A_tart_tepe", "1A_tart_uzama", "1A_tart_a_deg",
        "1B_kod_fft", "1B_kod_faxis", "1B_kod_tektar", "1B_graf_ucspek", "1B_tart_baskin", "1B_tart_filtre", "1B_tart_yxkars",
        "2A_kod_ornekle", "2A_kod_ayrconv", "2A_graf_xn", "2A_graf_hn", "2A_graf_kars", "2A_tart_genlik", "2A_tart_nyquist", "2A_tart_alias",
        "2B_kod_fft", "2B_kod_faxis", "2B_graf_spek", "2B_tart_benzer", "2B_tart_cozunur", "2B_tart_filtre",
        "SN_konv_kars", "SN_fft_kars", "SN_hf0_hesap",
    ]
    sub_headers = [("Sira", 5), ("Ogrenci No", 14), ("Ad Soyad", 18)] + [(k, 12) for k in score_keys]
    for col_idx, (h, w) in enumerate(sub_headers, 1):
        c = ws2.cell(row=1, column=col_idx, value=h)
        c.font = Font(name="Arial", bold=True, size=8, color="FFFFFF")
        c.fill = header_fill if col_idx <= 3 else sub_header_fill
        c.alignment = center
        c.border = thin_border
        ws2.column_dimensions[get_column_letter(col_idx)].width = w
    for i, r in enumerate(results, 1):
        ev = r["eval"]
        ws2.cell(row=i + 1, column=1, value=i)
        ws2.cell(row=i + 1, column=2, value=r["ogr_no"] or "?")
        ws2.cell(row=i + 1, column=3, value=r["ogr_name"] or "?")
        for j, key in enumerate(score_keys, start=4):
            val = ev["scores"].get(key, 0)
            conf = ev["confidences"].get(key, 0)
            c = ws2.cell(row=i + 1, column=j, value=val)
            c.alignment = center
            c.border = thin_border
            if val == 0:
                c.fill = bad_fill
            elif conf < 0.55:
                c.fill = warn_fill
        if i % 2 == 0:
            for col_idx in range(1, len(sub_headers) + 1):
                cell = ws2.cell(row=i + 1, column=col_idx)
                if cell.fill == PatternFill():
                    cell.fill = light_gray
    ws2.freeze_panes = "D2"

    # Sayfa 3 - kanit/gerekce
    ws3 = wb.create_sheet("Kanit")
    headers3 = [("Ogrenci No", 14), ("Kriter", 20), ("Puan", 8), ("Guven", 8), ("Gerekce", 80)]
    for col_idx, (h, w) in enumerate(headers3, 1):
        c = ws3.cell(row=1, column=col_idx, value=h)
        c.font = header_font
        c.fill = PatternFill("solid", fgColor="28A745")
        c.alignment = center
        c.border = thin_border
        ws3.column_dimensions[get_column_letter(col_idx)].width = w
    row_idx = 2
    for r in results:
        ev = r["eval"]
        for key in score_keys:
            ws3.cell(row=row_idx, column=1, value=r["ogr_no"] or "?")
            ws3.cell(row=row_idx, column=2, value=key)
            ws3.cell(row=row_idx, column=3, value=ev["scores"].get(key, 0))
            ws3.cell(row=row_idx, column=4, value=ev["confidences"].get(key, 0))
            ws3.cell(row=row_idx, column=5, value=ev["reasons"].get(key, ""))
            for col_idx in range(1, 6):
                ws3.cell(row=row_idx, column=col_idx).border = thin_border
            row_idx += 1
    ws3.freeze_panes = "A2"

    # Sayfa 4 - beklenen degerler
    ws4 = wb.create_sheet("Beklenen Degerler")
    exp_headers = [("N", 5), ("a", 5), ("f0", 5), ("T", 5), ("y_peak", 10), ("y_peak_t", 10), ("dominant_f", 10), ("H_f0", 10), ("-3dB", 10), ("zayif_%", 10), ("y_d_peak*Ts", 12), ("df_cont", 10), ("df_disc", 10), ("alias@0.1", 10)]
    for col_idx, (h, w) in enumerate(exp_headers, 1):
        c = ws4.cell(row=1, column=col_idx, value=h)
        c.font = header_font
        c.fill = PatternFill("solid", fgColor="6F42C1")
        c.alignment = center
        c.border = thin_border
        ws4.column_dimensions[get_column_letter(col_idx)].width = w
    for N in range(100):
        p = compute_params(N)
        e = compute_expected(p["a"], p["f0"], p["T"])
        vals = [N, p["a"], p["f0"], p["T"], e["y_peak"], e["y_peak_t"], e["dominant_f"], e["H_f0"], e["freq_3db"], e["attenuation_pct"], e["y_d_peak_scaled"], e["df_cont"], e["df_disc"], "EVET" if e["alias_if_ts_0_1"] else "HAYIR"]
        for col_idx, val in enumerate(vals, 1):
            c = ws4.cell(row=N + 2, column=col_idx, value=val)
            c.alignment = center
            c.border = thin_border
        if N % 2 == 0:
            for col_idx in range(1, len(exp_headers) + 1):
                ws4.cell(row=N + 2, column=col_idx).fill = light_gray
    ws4.freeze_panes = "A2"

    ws3 = wb.create_sheet("Gorsel Kontrol")
    vh = [("Sira",5),("Ogrenci No",14),("Ad Soyad",18),("GrafikSf",8),("ScreenshotSf",12),("TekrarPair",10),("Detay",70)]
    for col_idx, (h, w) in enumerate(vh, 1):
        c = ws3.cell(row=1, column=col_idx, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = thin_border
        ws3.column_dimensions[get_column_letter(col_idx)].width = w
    for i, r in enumerate(results, 1):
        ev = r["eval"]
        vi = ev.get("visual_info", {}) or {}
        details = vi.get("page_details", []) or []
        detail_txt = " | ".join([f"s{d['page']}:g={int(d['looks_graph'])},img={d['images']},dr={d['drawings']},e={d['edge_ratio']}" for d in details[:12]])
        row = [i, r["ogr_no"] or "?", r["ogr_name"] or "?", ev.get("graph_like_pages",0), ev.get("screenshot_like_pages",0), ev.get("repeated_graph_pairs",0), detail_txt]
        for col_idx, val in enumerate(row, 1):
            c = ws3.cell(row=i+1, column=col_idx, value=val)
            c.border = thin_border
            c.alignment = left_wrap if col_idx in (2,3,7) else center
        if ev.get("repeated_graph_pairs",0):
            ws3.cell(row=i+1, column=6).fill = warn_fill

    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Sinyaller ve Sistemler proje PDF degerlendirme v6")
    parser.add_argument("--extract-text-worker", nargs=2, metavar=("PDF", "OUT"), help=argparse.SUPPRESS)
    parser.add_argument("input_dir", nargs="?", default=None, help="PDF klasoru.")
    args = parser.parse_args()

    # Subprocess worker mode (called internally for pdfplumber timeout)
    if args.extract_text_worker:
        sys.exit(_run_extract_text_worker(*args.extract_text_worker))

    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidate_dirs = [os.path.abspath(args.input_dir)] if args.input_dir else [script_dir, os.getcwd()]

    pdf_files: List[str] = []
    chosen_dir = ""
    for d in candidate_dirs:
        found = sorted(glob.glob(os.path.join(d, "*.pdf")))
        if found:
            pdf_files = found
            chosen_dir = d
            break

    if not pdf_files:
        print(f"\nHATA: PDF bulunamadi. Denenen klasorler: {safe_console_text(candidate_dirs)}")
        print("Kullanim: python degerlendir_v6.py <pdf_klasoru>")
        sys.exit(1)

    print("\n" + "=" * 68)
    print("  SINYAL VE SISTEMLER - PROJE DEGERLENDIRME V6")
    print(f"  {len(pdf_files)} adet PDF bulundu")
    print(f"  Klasor: {safe_console_text(chosen_dir)}")
    print("=" * 68 + "\n")

    results = []
    unreadable = []

    def _empty_eval(note_key, note_val, ss=False):
        return {
            "criteria": {}, "scores": {}, "confidences": {}, "reasons": {},
            "notes": {note_key: note_val},
            "b1a": 0, "b1b": 0, "b2a": 0, "b2b": 0, "sonuc": 0,
            "ham_toplam": 0, "ceza_dosya_adi": -5, "ceza_etiket": 0, "ceza_sablon": 0,
            "template_uncleaned": False,
            "net_toplam": 0, "param_ok": False, "screenshot_uyari": ss,
            "overall_confidence": 0.0, "manual_review": True,
            "reported_params": {}, "visual_info": {},
            "graph_like_pages": 0, "screenshot_like_pages": 0, "repeated_graph_pairs": 0,
            "id_info": {},
        }

    for i, pdf_path in enumerate(pdf_files, 1):
        filename = os.path.basename(pdf_path)
        display = safe_console_text(filename)
        print(f"[{i:3d}/{len(pdf_files)}] {display[:50]:50s}", end=" ")

        try:
            text = extract_text(pdf_path)
        except KeyboardInterrupt:
            print("ATLANDI (kullanici iptal)")
            results.append({"ogr_no": None, "ogr_name": filename, "N": None,
                            "params": None, "eval": _empty_eval("HATA", "Kullanici iptal etti", True)})
            continue
        except Exception as e:
            print(f"HATA: {type(e).__name__}")
            results.append({"ogr_no": None, "ogr_name": filename, "N": None,
                            "params": None, "eval": _empty_eval("HATA", f"Islem hatasi: {e}", True)})
            continue
        if text is None or len(text) < 80:
            print("HATA: Metin cikarilmadi")
            unreadable.append(filename)
            results.append({"ogr_no": None, "ogr_name": filename, "N": None,
                            "params": None, "eval": _empty_eval("HATA", "PDF okunamadi", True)})
            continue

        ogr_no, ogr_name, N, id_info = extract_student_info(text, filename)
        if N is None:
            print("UYARI: Ogrenci No bulunamadi (metin icinde)")
            results.append({"ogr_no": ogr_no, "ogr_name": ogr_name or filename, "N": None,
                            "params": None, "eval": _empty_eval("UYARI", "Ogrenci No metin icinde bulunamadi")})
            continue

        params = compute_params(N)
        expected = compute_expected(params["a"], params["f0"], params["T"])
        try:
            visual_info = analyze_pdf_visual(pdf_path)
        except Exception:
            visual_info = {}
        ev = evaluate_pdf(text, filename, params, expected, visual_info=visual_info)

        # Kimlik notlari
        ev["id_info"] = id_info
        if id_info.get("src_no") == "filename_fallback":
            ev["notes"]["KIMLIK"] = "Numara metinden alinamadi, dosya adina dusuldu - manuel kontrol gerekli"
        if id_info.get("txt_no") and id_info.get("fn_no") and not id_info["fn_text_match"]:
            ev["notes"]["KIMLIK_FARK"] = (
                f"Metin: {id_info['txt_no']}, Dosya adi: {id_info['fn_no']} - uyusmuyor "
                f"(metin onceligi alindi)"
            )

        # Dosya yeniden adlandirma: kimlik tespit edildiyse {ogr_no}_{Ad}_Project.pdf
        # formatina cevir. LMS dosyalari sistemin verdigi adla geldigi icin
        # bu adim hocanin sonradan ogrenciyi bulmasini kolaylastirir.
        renamed_path = None
        if RENAME_FILES and ogr_no and ogr_name:
            new_path, err = rename_pdf_to_canonical(pdf_path, ogr_no, ogr_name)
            if err is None and new_path != pdf_path:
                renamed_path = new_path
                ev["notes"]["RENAME"] = f"Dosya yeniden adlandirildi: {os.path.basename(new_path)}"
            elif err and err != "kimlik eksik":
                ev["notes"]["RENAME_HATA"] = f"Yeniden adlandirilamadi: {err}"

        results.append({
            "ogr_no": ogr_no, "ogr_name": ogr_name, "N": N, "params": params, "eval": ev,
            "original_filename": filename,
            "current_filename": os.path.basename(renamed_path) if renamed_path else filename,
        })

        flags = []
        if not ev["param_ok"]:
            flags.append("PARAM")
        if ev.get("template_uncleaned"):
            flags.append("SABLON")
        if renamed_path:
            flags.append("RENAMED")
        if ev["manual_review"]:
            flags.append("MANUEL")
        suffix = f" ({', '.join(flags)})" if flags else ""
        print(f"N={N:2d} a={params['a']} f0={params['f0']} -> {ev['net_toplam']:3d}/100 guven={ev['overall_confidence']:.2f} grafikSf={ev.get('graph_like_pages',0)}{suffix}")

    output_path = os.path.join(chosen_dir, "notlar_v6.xlsx")
    create_excel(results, output_path)

    nets = [r["eval"]["net_toplam"] for r in results]
    confs = [r["eval"]["overall_confidence"] for r in results]
    manual_count = sum(1 for r in results if r["eval"].get("manual_review"))
    graph_counts = [r["eval"].get("graph_like_pages", 0) for r in results]

    print("\n" + "=" * 68)
    print("  TAMAMLANDI")
    print("=" * 68)
    print(f"  Toplam PDF      : {len(results)}")
    print(f"  Ortalama Not    : {sum(nets)/len(nets):.1f}" if nets else "  Ortalama Not    : -")
    print(f"  Medyan Not      : {median(nets):.1f}" if nets else "  Medyan Not      : -")
    print(f"  Ortalama Guven  : {sum(confs)/len(confs):.2f}" if confs else "  Ortalama Guven  : -")
    print(f"  Ort. GrafikSf   : {sum(graph_counts)/len(graph_counts):.1f}" if graph_counts else "  Ort. GrafikSf   : -")
    print(f"  Manuel Inceleme : {manual_count}")
    print(f"  Okunamayan      : {len(unreadable)}")
    print(f"  Excel dosyasi   : {safe_console_text(output_path)}")
    print("=" * 68 + "\n")


if __name__ == "__main__":
    main()