"""
Floor Time Schedule Generator untuk Brighton Real Estate
=========================================================

Aplikasi ini dibuat untuk membantu Staff Operasional membuat jadwal Floor Time
per bulan, lalu memecahnya menjadi jadwal per minggu yang rapi, adil, dan bisa
langsung dicetak.

Fitur utama:
1. Jumlah agen fleksibel. Agen bisa sedikit atau banyak, aplikasi akan menyesuaikan.
2. Jadwal digenerate per bulan dan otomatis dikelompokkan per minggu.
3. Tanggal merah atau hari kantor tutup bisa dipilih sebelum generate.
4. Setiap agen hanya mendapat 1 jadwal dalam 1 minggu.
5. Jadwal diacak secara menyeluruh setiap bulan.
6. Request urgent agen bisa dimasukkan untuk tanggal dan shift tertentu.
7. Hasil bisa dilihat di website, diunduh ke Excel, gambar PNG, dan PDF.

Cara menjalankan:
    pip install -r requirements.txt
    streamlit run app.py

Catatan teknis:
- Aplikasi ini sengaja dibuat dalam satu file agar mudah dipindahkan ke laptop lain.
- Seluruh fungsi penting diberi dokumentasi agar mudah dikembangkan kembali.
"""

from __future__ import annotations

import calendar
import html
import io
import os
import random
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from math import ceil
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont


# -----------------------------------------------------------------------------
# KONFIGURASI DASAR
# -----------------------------------------------------------------------------

APP_TITLE = "Brighton Floor Time Generator"
DEFAULT_HUB_NAME = "HUB CIBUBUR"
DEFAULT_LOGO_TEXT = "Brighton"

# Nama bulan dan hari memakai format Indonesia agar output siap dipakai.
MONTH_NAMES_ID = {
    1: "Januari",
    2: "Februari",
    3: "Maret",
    4: "April",
    5: "Mei",
    6: "Juni",
    7: "Juli",
    8: "Agustus",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Desember",
}

DAY_NAMES_ID = {
    0: "SENIN",
    1: "SELASA",
    2: "RABU",
    3: "KAMIS",
    4: "JUMAT",
    5: "SABTU",
    6: "MINGGU",
}

# Shift standar mengikuti pola pada contoh gambar.
# Senin sampai Jumat: 08.00-12.30 dan 12.30-17.00
# Sabtu: 08.00-11.30 dan 11.30-15.00
SHIFT_TEMPLATES = {
    0: [("Pagi", "08.00-12.30"), ("Siang", "12.30-17.00")],
    1: [("Pagi", "08.00-12.30"), ("Siang", "12.30-17.00")],
    2: [("Pagi", "08.00-12.30"), ("Siang", "12.30-17.00")],
    3: [("Pagi", "08.00-12.30"), ("Siang", "12.30-17.00")],
    4: [("Pagi", "08.00-12.30"), ("Siang", "12.30-17.00")],
    5: [("Pagi", "08.00-11.30"), ("Siang", "11.30-15.00")],
}

SHIFT_NAME_TO_INDEX = {"Pagi": 0, "Siang": 1}

# Skala render untuk output gambar/PDF agar hasil unduhan lebih tajam.
EXPORT_RENDER_SCALE = 2


# -----------------------------------------------------------------------------
# DATA CLASS
# -----------------------------------------------------------------------------

@dataclass(frozen=True)
class ShiftSlot:
    """Representasi satu slot jadwal.

    Satu slot berarti gabungan antara minggu, tanggal, hari, dan shift.
    Contoh: Minggu 1, Selasa 02 Juni 2026, shift Pagi 08.00-12.30.
    """

    week_no: int
    tanggal: date
    day_name: str
    shift_index: int
    shift_name: str
    time_label: str
    base_capacity: int

    @property
    def key(self) -> str:
        """Key unik untuk menyimpan daftar agen pada slot ini."""
        return f"{self.tanggal.isoformat()}__{self.shift_index}"


@dataclass
class UrgentRequest:
    """Request urgent dari agen untuk ditempatkan pada tanggal dan shift tertentu."""

    agent: str
    tanggal: date
    shift_index: int
    note: str = ""


# -----------------------------------------------------------------------------
# UTILITAS FORMAT
# -----------------------------------------------------------------------------

def format_date_id(d: date, include_year: bool = False) -> str:
    """Format tanggal ke bahasa Indonesia.

    Args:
        d: Objek tanggal.
        include_year: Bila True, tahun akan ditampilkan.

    Returns:
        Contoh tanpa tahun: "02 Juni"
        Contoh dengan tahun: "02 Juni 2026"
    """
    if include_year:
        return f"{d.day:02d} {MONTH_NAMES_ID[d.month]} {d.year}"
    return f"{d.day:02d} {MONTH_NAMES_ID[d.month]}"


def date_option_label(d: date) -> str:
    """Label tanggal untuk selectbox dan multiselect."""
    return f"{DAY_NAMES_ID[d.weekday()]} | {format_date_id(d, include_year=True)}"


def week_range_label(week_dates: List[date]) -> str:
    """Label rentang tanggal untuk judul poster jadwal mingguan."""
    if not week_dates:
        return ""
    start = min(week_dates)
    end = max(week_dates)
    if start.month == end.month and start.year == end.year:
        return f"{start.day:02d} {MONTH_NAMES_ID[start.month].upper()} - {end.day:02d} {MONTH_NAMES_ID[end.month].upper()} {end.year}"
    return f"{format_date_id(start, True).upper()} - {format_date_id(end, True).upper()}"


def clean_agent_name(value: object) -> str:
    """Membersihkan nama agen dari spasi berlebih dan bullet.

    Nama sengaja tidak diubah total menjadi kapital agar kode atau format asli
    dari operasional tetap bisa dipertahankan.
    """
    text = "" if value is None else str(value)
    text = text.strip().strip("-•*;")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def html_escape(text: object) -> str:
    """Escape teks sebelum dimasukkan ke HTML agar aman ditampilkan."""
    return html.escape("" if text is None else str(text))


# -----------------------------------------------------------------------------
# PARSING INPUT AGEN
# -----------------------------------------------------------------------------

def parse_agents_from_text(raw_text: str) -> Tuple[List[str], List[str]]:
    """Mengubah input textarea menjadi daftar agen unik.

    Format input yang didukung:
        PAULA (DUCC)
        AZWAR (IBEX)
        MONIKA (NDEP)

    Args:
        raw_text: Teks daftar agen, satu nama per baris.

    Returns:
        Tuple berisi:
        - list agen unik sesuai urutan input
        - list duplikat yang ditemukan
    """
    agents: List[str] = []
    duplicates: List[str] = []
    seen = set()

    for line in raw_text.splitlines():
        name = clean_agent_name(line)
        if not name:
            continue
        key = name.casefold()
        if key in seen:
            duplicates.append(name)
            continue
        seen.add(key)
        agents.append(name)

    return agents, duplicates


def read_agents_from_upload(uploaded_file) -> List[str]:
    """Membaca file CSV atau Excel yang berisi daftar agen.

    Aturan pembacaan:
    - Jika ada kolom bernama 'Agen', 'Nama', atau 'Nama Agen', kolom itu dipakai.
    - Jika tidak ada, kolom pertama akan dipakai.
    - Baris kosong otomatis diabaikan.
    """
    if uploaded_file is None:
        return []

    filename = uploaded_file.name.lower()
    if filename.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    elif filename.endswith((".xlsx", ".xls")):
        df = pd.read_excel(uploaded_file)
    else:
        st.warning("Format file belum didukung. Gunakan CSV atau Excel.")
        return []

    if df.empty:
        return []

    preferred_columns = ["Agen", "Nama", "Nama Agen", "agent", "nama", "nama agen"]
    selected_column = None
    for col in df.columns:
        if str(col).strip() in preferred_columns:
            selected_column = col
            break
    if selected_column is None:
        selected_column = df.columns[0]

    agents = []
    seen = set()
    for value in df[selected_column].tolist():
        name = clean_agent_name(value)
        if not name:
            continue
        key = name.casefold()
        if key not in seen:
            seen.add(key)
            agents.append(name)
    return agents


# -----------------------------------------------------------------------------
# PEMBENTUKAN KALENDER BULANAN
# -----------------------------------------------------------------------------

def get_month_work_dates(year: int, month: int) -> List[date]:
    """Mengambil tanggal kerja untuk jadwal bulanan dengan minggu utuh.

    Prinsip operasional yang dipakai:
    - Minggu berjalan dari Senin sampai Sabtu.
    - Minggu tidak dijadwalkan.
    - Jika akhir bulan jatuh di tengah minggu, jadwal otomatis diteruskan
      sampai Sabtu pada bulan berikutnya agar poster terakhir tetap 1 minggu utuh.

    Contoh:
        Juni 2026 berakhir pada Selasa, sehingga minggu terakhir akan berisi
        Senin 29 Juni sampai Sabtu 04 Juli 2026. Ini mencegah jadwal minggu
        terakhir tampil menggantung hanya 1-2 hari.
    """
    first_day = date(year, month, 1)
    last_day = date(year, month, calendar.monthrange(year, month)[1])

    # Cari tanggal aktif terakhir di bulan tersebut. Jika akhir bulan jatuh hari
    # Minggu, tanggal aktif terakhirnya adalah Sabtu sebelum hari Minggu itu.
    last_active = last_day
    while last_active.weekday() == 6:
        last_active -= timedelta(days=1)

    # Perpanjang sampai Sabtu pada minggu yang sama.
    end_saturday = last_active + timedelta(days=(5 - last_active.weekday()))

    dates: List[date] = []
    current = first_day
    while current <= end_saturday:
        if current.weekday() <= 5:
            dates.append(current)
        current += timedelta(days=1)
    return dates


def group_dates_by_calendar_week(work_dates: Iterable[date]) -> Dict[int, List[date]]:
    """Mengelompokkan tanggal kerja berdasarkan minggu kalender.

    Minggu dimulai dari Senin. Bila awal atau akhir bulan tidak penuh, aplikasi
    tetap akan membuat grup minggu parsial agar tidak ada tanggal yang hilang.
    """
    grouped_by_monday: Dict[date, List[date]] = {}
    for d in work_dates:
        monday = d - timedelta(days=d.weekday())
        grouped_by_monday.setdefault(monday, []).append(d)

    result: Dict[int, List[date]] = {}
    for idx, monday in enumerate(sorted(grouped_by_monday.keys()), start=1):
        result[idx] = sorted(grouped_by_monday[monday])
    return result


def build_slots_by_week(
    weeks: Dict[int, List[date]],
    closed_dates: Iterable[date],
    weekday_capacity: int,
    saturday_capacity: int,
) -> Dict[int, List[ShiftSlot]]:
    """Membuat slot shift untuk setiap minggu.

    Args:
        weeks: Mapping nomor minggu ke daftar tanggal.
        closed_dates: Tanggal merah atau tanggal kantor tutup.
        weekday_capacity: Kapasitas awal per shift untuk Senin sampai Jumat.
        saturday_capacity: Kapasitas awal per shift untuk Sabtu.

    Returns:
        Mapping nomor minggu ke daftar ShiftSlot.
    """
    closed_set = set(closed_dates)
    slots_by_week: Dict[int, List[ShiftSlot]] = {}

    for week_no, dates_in_week in weeks.items():
        slots: List[ShiftSlot] = []
        for d in dates_in_week:
            if d in closed_set:
                continue
            if d.weekday() == 6:
                continue

            base_capacity = saturday_capacity if d.weekday() == 5 else weekday_capacity
            for shift_index, (shift_name, time_label) in enumerate(SHIFT_TEMPLATES[d.weekday()]):
                slots.append(
                    ShiftSlot(
                        week_no=week_no,
                        tanggal=d,
                        day_name=DAY_NAMES_ID[d.weekday()],
                        shift_index=shift_index,
                        shift_name=shift_name,
                        time_label=time_label,
                        base_capacity=base_capacity,
                    )
                )
        slots_by_week[week_no] = slots

    return slots_by_week


# -----------------------------------------------------------------------------
# REQUEST URGENT
# -----------------------------------------------------------------------------

def parse_urgent_requests(
    urgent_df: pd.DataFrame,
    label_to_date: Dict[str, date],
    agent_list: List[str],
) -> Tuple[List[UrgentRequest], List[str]]:
    """Memvalidasi dan mengubah tabel request urgent menjadi objek UrgentRequest.

    Validasi awal meliputi:
    - Agen harus ada di daftar agen resmi.
    - Tanggal harus dipilih dari daftar tanggal bulan tersebut.
    - Shift wajib Pagi atau Siang.
    """
    requests: List[UrgentRequest] = []
    errors: List[str] = []
    agent_key_to_real = {agent.casefold(): agent for agent in agent_list}

    if urgent_df is None or urgent_df.empty:
        return requests, errors

    for idx, row in urgent_df.iterrows():
        raw_agent = clean_agent_name(row.get("Agen", ""))
        raw_date_label = clean_agent_name(row.get("Tanggal", ""))
        raw_shift = clean_agent_name(row.get("Shift", ""))
        raw_note = clean_agent_name(row.get("Catatan", ""))

        # Baris kosong dari data_editor tidak perlu diproses.
        if not raw_agent and not raw_date_label and not raw_shift and not raw_note:
            continue

        row_no = idx + 1
        if not raw_agent:
            errors.append(f"Request urgent baris {row_no}: nama agen belum diisi.")
            continue
        if raw_agent.casefold() not in agent_key_to_real:
            errors.append(f"Request urgent baris {row_no}: agen '{raw_agent}' tidak ada di daftar agen.")
            continue
        if raw_date_label not in label_to_date:
            errors.append(f"Request urgent baris {row_no}: tanggal belum valid atau belum dipilih.")
            continue
        if raw_shift not in SHIFT_NAME_TO_INDEX:
            errors.append(f"Request urgent baris {row_no}: shift harus Pagi atau Siang.")
            continue

        requests.append(
            UrgentRequest(
                agent=agent_key_to_real[raw_agent.casefold()],
                tanggal=label_to_date[raw_date_label],
                shift_index=SHIFT_NAME_TO_INDEX[raw_shift],
                note=raw_note,
            )
        )

    return requests, errors


# -----------------------------------------------------------------------------
# MESIN GENERATE JADWAL
# -----------------------------------------------------------------------------

def make_rng(seed_text: Optional[str]) -> random.Random:
    """Membuat random generator.

    Bila user mengisi kode audit, hasil acak bisa direproduksi kembali.
    Bila kosong, aplikasi memakai SystemRandom agar hasil benar-benar segar
    setiap generate dilakukan.
    """
    if seed_text and seed_text.strip():
        return random.Random(seed_text.strip())
    return random.SystemRandom()


def generate_schedule(
    agents: List[str],
    slots_by_week: Dict[int, List[ShiftSlot]],
    urgent_requests: List[UrgentRequest],
    auto_expand_capacity: bool,
    seed_text: Optional[str],
) -> Tuple[pd.DataFrame, Dict[int, Dict[str, List[Tuple[str, bool]]]], Dict[int, Dict[str, int]], List[str]]:
    """Generate jadwal floor time bulanan.

    Prinsip utama algoritma:
    1. Setiap minggu diproses terpisah.
    2. Request urgent ditempatkan lebih dahulu.
    3. Agen yang belum punya jadwal di minggu tersebut akan diacak.
    4. Setiap agen hanya boleh muncul satu kali per minggu.
    5. Slot dipilih secara seimbang berdasarkan load paling rendah.
    6. Bila jumlah agen melebihi kapasitas awal, kapasitas bisa dinaikkan otomatis.

    Returns:
        assignments_df:
            Data detail jadwal per baris agen.
        schedule:
            Mapping week_no -> slot_key -> list tuple (agent, is_urgent).
        capacity_by_week:
            Mapping week_no -> slot_key -> kapasitas final.
        warnings:
            Informasi penting yang perlu diketahui operator.
    """
    rng = make_rng(seed_text)
    warnings: List[str] = []
    rows: List[dict] = []
    schedule: Dict[int, Dict[str, List[Tuple[str, bool]]]] = {}
    capacity_by_week: Dict[int, Dict[str, int]] = {}

    # Index untuk mencari minggu berdasarkan tanggal.
    date_to_week: Dict[date, int] = {}
    slot_lookup_by_week: Dict[int, Dict[Tuple[date, int], ShiftSlot]] = {}
    for week_no, slots in slots_by_week.items():
        slot_lookup_by_week[week_no] = {}
        for slot in slots:
            date_to_week[slot.tanggal] = week_no
            slot_lookup_by_week[week_no][(slot.tanggal, slot.shift_index)] = slot

    # Group request urgent berdasarkan minggu.
    urgent_by_week: Dict[int, List[UrgentRequest]] = {}
    for req in urgent_requests:
        week_no = date_to_week.get(req.tanggal)
        if week_no is None:
            warnings.append(
                f"Request urgent {req.agent} pada {format_date_id(req.tanggal, True)} dilewati karena tanggal tersebut libur atau tidak punya slot."
            )
            continue
        urgent_by_week.setdefault(week_no, []).append(req)

    for week_no, slots in slots_by_week.items():
        schedule[week_no] = {slot.key: [] for slot in slots}
        capacity_by_week[week_no] = {slot.key: slot.base_capacity for slot in slots}

        if not slots:
            warnings.append(f"Minggu {week_no} tidak memiliki hari aktif karena seluruh tanggalnya libur atau di luar hari kerja.")
            continue

        # Jika jumlah agen lebih besar dari kapasitas total awal, kapasitas dinaikkan
        # merata agar seluruh agen tetap mendapat 1 jadwal dalam minggu tersebut.
        total_initial_capacity = sum(capacity_by_week[week_no].values())
        if len(agents) > total_initial_capacity:
            if auto_expand_capacity:
                minimum_capacity_per_slot = ceil(len(agents) / len(slots))
                for slot in slots:
                    capacity_by_week[week_no][slot.key] = max(
                        capacity_by_week[week_no][slot.key], minimum_capacity_per_slot
                    )
                warnings.append(
                    f"Minggu {week_no}: kapasitas per shift dinaikkan otomatis karena jumlah agen ({len(agents)}) lebih besar dari kapasitas awal ({total_initial_capacity})."
                )
            else:
                warnings.append(
                    f"Minggu {week_no}: jumlah agen ({len(agents)}) melebihi kapasitas awal ({total_initial_capacity}). Aktifkan 'naikkan kapasitas otomatis' agar semua agen masuk."
                )

        used_agents = set()
        used_agent_names_for_warning = set()

        # 1) Tempatkan request urgent lebih dahulu.
        for req in urgent_by_week.get(week_no, []):
            slot = slot_lookup_by_week[week_no].get((req.tanggal, req.shift_index))
            if slot is None:
                warnings.append(
                    f"Request urgent {req.agent} pada {format_date_id(req.tanggal, True)} shift {req.shift_index + 1} dilewati karena slot tidak tersedia."
                )
                continue

            if req.agent in used_agents:
                warnings.append(
                    f"Minggu {week_no}: request urgent {req.agent} dilewati karena agen tersebut sudah punya jadwal di minggu yang sama."
                )
                continue

            if len(schedule[week_no][slot.key]) >= capacity_by_week[week_no][slot.key]:
                if auto_expand_capacity:
                    capacity_by_week[week_no][slot.key] += 1
                    warnings.append(
                        f"Minggu {week_no}: kapasitas slot {slot.day_name} {format_date_id(slot.tanggal)} {slot.time_label} dinaikkan karena request urgent."
                    )
                else:
                    warnings.append(
                        f"Minggu {week_no}: request urgent {req.agent} tidak masuk karena slot penuh."
                    )
                    continue

            schedule[week_no][slot.key].append((req.agent, True))
            used_agents.add(req.agent)
            used_agent_names_for_warning.add(req.agent)

        # 2) Acak seluruh agen yang belum mendapat slot pada minggu ini.
        remaining_agents = [agent for agent in agents if agent not in used_agents]
        rng.shuffle(remaining_agents)

        for agent in remaining_agents:
            # Pilih slot yang belum penuh.
            available_slots = [
                slot for slot in slots
                if len(schedule[week_no][slot.key]) < capacity_by_week[week_no][slot.key]
            ]

            # Bila semua penuh tetapi auto-expand aktif, pilih slot dengan load paling rendah.
            if not available_slots:
                if auto_expand_capacity:
                    least_loaded_slot = min(slots, key=lambda s: len(schedule[week_no][s.key]))
                    capacity_by_week[week_no][least_loaded_slot.key] += 1
                    available_slots = [least_loaded_slot]
                else:
                    warnings.append(
                        f"Minggu {week_no}: {agent} belum mendapat jadwal karena semua slot penuh."
                    )
                    continue

            # Seimbangkan isi slot berdasarkan rasio load terhadap kapasitas.
            min_ratio = min(
                len(schedule[week_no][slot.key]) / max(capacity_by_week[week_no][slot.key], 1)
                for slot in available_slots
            )
            balanced_candidates = [
                slot for slot in available_slots
                if len(schedule[week_no][slot.key]) / max(capacity_by_week[week_no][slot.key], 1) == min_ratio
            ]
            selected_slot = rng.choice(balanced_candidates)
            schedule[week_no][selected_slot.key].append((agent, False))
            used_agents.add(agent)

        # 3) Simpan hasil ke dataframe detail.
        for slot in slots:
            for order_no, (agent, is_urgent) in enumerate(schedule[week_no][slot.key], start=1):
                rows.append(
                    {
                        "Minggu": week_no,
                        "Tanggal": slot.tanggal.isoformat(),
                        "Hari": slot.day_name,
                        "Shift": slot.shift_name,
                        "Jam": slot.time_label,
                        "Urutan": order_no,
                        "Agen": agent,
                        "Urgent": "Ya" if is_urgent else "Tidak",
                    }
                )

    assignments_df = pd.DataFrame(rows)
    return assignments_df, schedule, capacity_by_week, warnings


def validate_schedule(assignments_df: pd.DataFrame, agents: List[str], active_weeks: Iterable[int]) -> Tuple[bool, List[str]]:
    """Memeriksa kualitas hasil generate.

    Validasi yang dilakukan:
    - Setiap agen harus muncul tepat 1 kali pada setiap minggu aktif.
    - Tidak boleh ada agen yang double pada minggu yang sama.
    """
    messages: List[str] = []
    if assignments_df.empty:
        return False, ["Jadwal kosong. Pastikan daftar agen dan hari aktif sudah benar."]

    valid = True
    for week_no in active_weeks:
        week_df = assignments_df[assignments_df["Minggu"] == week_no]
        if week_df.empty:
            continue

        counts = week_df.groupby("Agen").size().to_dict()
        missing_agents = [agent for agent in agents if counts.get(agent, 0) == 0]
        duplicated_agents = [agent for agent, count in counts.items() if count > 1]

        if missing_agents:
            valid = False
            messages.append(f"Minggu {week_no}: agen belum mendapat jadwal: {', '.join(missing_agents)}")
        if duplicated_agents:
            valid = False
            messages.append(f"Minggu {week_no}: agen mendapat jadwal lebih dari sekali: {', '.join(duplicated_agents)}")

    if valid:
        messages.append("Validasi berhasil: setiap agen mendapat tepat 1 jadwal pada setiap minggu aktif.")
    return valid, messages


# -----------------------------------------------------------------------------
# OUTPUT HTML SIAP CETAK
# -----------------------------------------------------------------------------

def build_schedule_css() -> str:
    """CSS poster jadwal dengan nuansa Brighton."""
    return """
    <style>
        :root {
            --brighton-yellow: #ffd10a;
            --brighton-black: #111111;
            --soft-gray: #f3f3f3;
            --line-gray: #d9d9d9;
        }
        * { box-sizing: border-box; }
        body {
            margin: 0;
            font-family: Arial, Helvetica, sans-serif;
            color: var(--brighton-black);
            background: #ffffff;
        }
        .poster {
            width: 100%;
            max-width: 980px;
            margin: 0 auto 28px auto;
            background: #ffffff;
            border: 1px solid #efefef;
            page-break-after: always;
        }
        .brand-bar {
            position: relative;
            background: var(--brighton-yellow);
            height: 96px;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 0 24px;
        }
        .brand-name {
            font-size: 58px;
            font-weight: 800;
            letter-spacing: -2px;
            line-height: 1;
        }
        .brand-name .white-o { color: #ffffff; }
        .hub-name {
            position: absolute;
            right: 26px;
            bottom: 14px;
            color: #ffffff;
            font-size: 21px;
            font-weight: 800;
        }
        .title-wrap {
            text-align: center;
            padding: 22px 30px 14px 30px;
        }
        .schedule-title {
            font-size: 42px;
            line-height: 1;
            margin: 0;
            font-weight: 900;
            letter-spacing: 1px;
        }
        .thin-line {
            width: 58%;
            height: 2px;
            background: #f2dc76;
            margin: 12px auto 7px auto;
        }
        .date-range {
            font-size: 24px;
            font-weight: 700;
            letter-spacing: 1px;
        }
        .week-label {
            font-size: 16px;
            margin-top: 6px;
            color: #555555;
            font-weight: 700;
        }
        .schedule-body {
            padding: 12px 54px 34px 54px;
        }
        .day-section {
            margin-bottom: 20px;
        }
        .day-header {
            display: flex;
            align-items: end;
            gap: 6px;
            border-bottom: 4px solid var(--brighton-yellow);
            margin-bottom: 8px;
        }
        .day-badge {
            min-width: 168px;
            background: var(--brighton-yellow);
            padding: 10px 18px 12px 18px;
            font-size: 30px;
            font-weight: 900;
            text-align: center;
            clip-path: polygon(0 0, 100% 0, 95% 100%, 0% 100%);
        }
        .date-label {
            font-size: 22px;
            font-weight: 800;
            padding-bottom: 5px;
        }
        .shift-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 32px;
            padding-left: 34px;
        }
        .shift-box {
            display: grid;
            grid-template-columns: 168px 1fr;
            gap: 14px;
            align-items: start;
        }
        .time-label {
            font-size: 18px;
            font-weight: 800;
            color: #4a4a4a;
            border-bottom: 2px solid var(--line-gray);
            padding: 5px 0 8px 0;
            white-space: nowrap;
        }
        .clock {
            display: inline-flex;
            width: 24px;
            height: 24px;
            border: 2px solid var(--brighton-yellow);
            border-radius: 50%;
            align-items: center;
            justify-content: center;
            margin-right: 8px;
            vertical-align: -3px;
            font-size: 13px;
            color: var(--brighton-yellow);
        }
        .agent-list {
            list-style: none;
            margin: 0;
            padding: 0;
        }
        .agent-list li {
            position: relative;
            font-size: 23px;
            line-height: 1.28;
            font-weight: 900;
            padding: 3px 0 5px 24px;
            border-bottom: 2px solid var(--line-gray);
            min-height: 34px;
        }
        .agent-list li::before {
            content: "•";
            position: absolute;
            left: 0;
            top: 1px;
            color: #222222;
            font-size: 28px;
            line-height: 1;
        }
        .agent-list .urgent-tag {
            display: inline-block;
            margin-left: 6px;
            font-size: 10px;
            background: var(--brighton-yellow);
            padding: 2px 5px;
            border-radius: 10px;
            vertical-align: 4px;
            font-weight: 900;
        }
        .special-hours {
            display: inline-block;
            margin-left: 12px;
            background: var(--brighton-yellow);
            color: #ffffff;
            padding: 6px 14px;
            font-size: 14px;
            font-weight: 900;
        }
        .notes {
            margin: 56px 34px 8px 34px;
            background: var(--soft-gray);
            padding: 24px 20px 18px 20px;
            position: relative;
        }
        .note-badge {
            position: absolute;
            left: 0;
            top: -38px;
            min-width: 198px;
            background: var(--brighton-yellow);
            padding: 10px 16px;
            font-size: 21px;
            font-weight: 900;
            clip-path: polygon(0 0, 100% 0, 97% 100%, 0% 100%);
        }
        .notes p {
            margin: 9px 0;
            font-size: 18px;
            line-height: 1.35;
        }
        .notes p::before {
            content: "•";
            font-weight: 900;
            margin-right: 8px;
        }
        .footer-bar {
            height: 16px;
            background: var(--brighton-yellow);
            border-top: 4px solid #ffffff;
        }
        .empty-day {
            color: #777777;
            font-size: 16px;
            margin-left: 36px;
            padding: 8px 0;
        }
        @media print {
            body { background: #ffffff; }
            .poster { border: none; margin: 0; max-width: 100%; }
        }
        @media (max-width: 820px) {
            .brand-name { font-size: 40px; }
            .hub-name { font-size: 14px; right: 16px; }
            .schedule-title { font-size: 28px; }
            .schedule-body { padding: 12px 18px 30px 18px; }
            .shift-grid { grid-template-columns: 1fr; padding-left: 0; gap: 14px; }
            .shift-box { grid-template-columns: 132px 1fr; }
            .agent-list li { font-size: 18px; }
            .day-badge { min-width: 130px; font-size: 22px; }
        }
    </style>
    """


def render_agent_list(agents: List[Tuple[str, bool]]) -> str:
    """Render daftar agen dalam satu shift."""
    if not agents:
        return "<ul class='agent-list'><li>&nbsp;</li></ul>"

    items = []
    for agent, _is_urgent in agents:
        # Status urgent sengaja tidak ditampilkan pada jadwal.
        # Request tetap diprioritaskan oleh sistem, tetapi labelnya hanya menjadi arsip internal admin.
        items.append(f"<li>{html_escape(agent)}</li>")
    return "<ul class='agent-list'>" + "".join(items) + "</ul>"


def render_week_html(
    week_no: int,
    week_dates: List[date],
    slots: List[ShiftSlot],
    schedule_for_week: Dict[str, List[Tuple[str, bool]]],
    hub_name: str,
    logo_text: str,
    notes: List[str],
) -> str:
    """Membuat HTML poster untuk satu minggu."""
    dates_in_slots = sorted({slot.tanggal for slot in slots})
    date_to_slots: Dict[date, List[ShiftSlot]] = {}
    for slot in slots:
        date_to_slots.setdefault(slot.tanggal, []).append(slot)

    day_sections = []
    for d in dates_in_slots:
        day_name = DAY_NAMES_ID[d.weekday()]
        special_hours = "<span class='special-hours'>Jam Operasional Khusus</span>" if d.weekday() == 5 else ""
        day_html = f"""
        <section class="day-section">
            <div class="day-header">
                <div class="day-badge">{html_escape(day_name)}</div>
                <div class="date-label">{html_escape(format_date_id(d))}</div>
                {special_hours}
            </div>
            <div class="shift-grid">
        """

        for slot in sorted(date_to_slots[d], key=lambda s: s.shift_index):
            agents = schedule_for_week.get(slot.key, [])
            day_html += f"""
                <div class="shift-box">
                    <div class="time-label"><span class="clock">◷</span>{html_escape(slot.time_label)}</div>
                    <div>{render_agent_list(agents)}</div>
                </div>
            """

        day_html += """
            </div>
        </section>
        """
        day_sections.append(day_html)

    if not day_sections:
        day_sections.append("<div class='empty-day'>Tidak ada jadwal aktif pada minggu ini.</div>")

    notes_html = "".join(f"<p>{html_escape(note)}</p>" for note in notes if note.strip())
    if not notes_html:
        notes_html = "<p>Agen yang mendapatkan jadwal floor time masih berada di kantor.</p>"

    # Logo Brighton diberi aksen putih pada huruf O agar sesuai identitas visual.
    safe_logo = html_escape(logo_text)
    if safe_logo.lower() == "brighton":
        safe_logo = "Bright<span class='white-o'>o</span>n"

    return f"""
    <div class="poster">
        <div class="brand-bar">
            <div class="brand-name">{safe_logo}</div>
            <div class="hub-name">{html_escape(hub_name)}</div>
        </div>
        <div class="title-wrap">
            <h1 class="schedule-title">FLOOR TIME SCHEDULE</h1>
            <div class="thin-line"></div>
            <div class="date-range">{html_escape(week_range_label(week_dates))}</div>
            <div class="week-label">MINGGU {week_no}</div>
        </div>
        <div class="schedule-body">
            {''.join(day_sections)}
            <div class="notes">
                <div class="note-badge">CATATAN:</div>
                {notes_html}
            </div>
        </div>
        <div class="footer-bar"></div>
    </div>
    """


def build_full_html_document(
    weeks: Dict[int, List[date]],
    slots_by_week: Dict[int, List[ShiftSlot]],
    schedule: Dict[int, Dict[str, List[Tuple[str, bool]]]],
    hub_name: str,
    logo_text: str,
    notes: List[str],
) -> str:
    """Membuat HTML lengkap untuk seluruh bulan."""
    week_pages = []
    for week_no, week_dates in weeks.items():
        week_pages.append(
            render_week_html(
                week_no=week_no,
                week_dates=week_dates,
                slots=slots_by_week.get(week_no, []),
                schedule_for_week=schedule.get(week_no, {}),
                hub_name=hub_name,
                logo_text=logo_text,
                notes=notes,
            )
        )

    return f"""
    <!doctype html>
    <html lang="id">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Floor Time Schedule</title>
        {build_schedule_css()}
    </head>
    <body>
        {''.join(week_pages)}
    </body>
    </html>
    """


# -----------------------------------------------------------------------------
# EXPORT EXCEL DAN CSV
# -----------------------------------------------------------------------------

def build_excel_file(
    assignments_df: pd.DataFrame,
    weeks: Dict[int, List[date]],
    slots_by_week: Dict[int, List[ShiftSlot]],
    schedule: Dict[int, Dict[str, List[Tuple[str, bool]]]],
    hub_name: str,
) -> bytes:
    """Membuat file Excel dengan sheet per minggu dan detail data mentah."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    yellow = "FFD10A"
    black = "111111"
    gray = "F3F3F3"
    white = "FFFFFF"
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:H1")
    ws["A1"] = f"FLOOR TIME SCHEDULE - {hub_name}"
    ws["A1"].font = Font(bold=True, size=16, color=black)
    ws["A1"].fill = PatternFill("solid", fgColor=yellow)
    ws["A1"].alignment = Alignment(horizontal="center")

    summary_headers = ["Minggu", "Rentang Tanggal", "Jumlah Agen Terjadwal", "Jumlah Slot Aktif"]
    for col_idx, header in enumerate(summary_headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor=gray)
        cell.border = border

    for row_idx, (week_no, dates_in_week) in enumerate(weeks.items(), start=4):
        week_df = assignments_df[assignments_df["Minggu"] == week_no] if not assignments_df.empty else pd.DataFrame()
        values = [
            week_no,
            week_range_label(dates_in_week),
            len(week_df),
            len(slots_by_week.get(week_no, [])),
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border

    for col_idx in range(1, 5):
        ws.column_dimensions[get_column_letter(col_idx)].width = 24


    # Sheet per minggu dengan format yang mudah dibaca.
    for week_no, dates_in_week in weeks.items():
        sheet_name = f"Minggu {week_no}"
        ws_week = wb.create_sheet(title=sheet_name[:31])
        ws_week.merge_cells("A1:E1")
        ws_week["A1"] = f"FLOOR TIME SCHEDULE - MINGGU {week_no}"
        ws_week["A1"].font = Font(bold=True, size=15, color=black)
        ws_week["A1"].fill = PatternFill("solid", fgColor=yellow)
        ws_week["A1"].alignment = Alignment(horizontal="center")

        ws_week.merge_cells("A2:E2")
        ws_week["A2"] = week_range_label(dates_in_week)
        ws_week["A2"].font = Font(bold=True, size=12)
        ws_week["A2"].alignment = Alignment(horizontal="center")

        headers = ["Hari", "Tanggal", "Shift", "Jam", "Agen"]
        for col_idx, header in enumerate(headers, start=1):
            cell = ws_week.cell(row=4, column=col_idx, value=header)
            cell.font = Font(bold=True, color=white)
            cell.fill = PatternFill("solid", fgColor=black)
            cell.border = border
            cell.alignment = Alignment(horizontal="center")

        row_idx = 5
        for slot in slots_by_week.get(week_no, []):
            agents_in_slot = schedule.get(week_no, {}).get(slot.key, [])
            agent_text = "\n".join(
                f"• {agent}"
                for agent, _is_urgent in agents_in_slot
            )
            values = [slot.day_name, format_date_id(slot.tanggal), slot.shift_name, slot.time_label, agent_text]
            for col_idx, value in enumerate(values, start=1):
                cell = ws_week.cell(row=row_idx, column=col_idx, value=value)
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            row_height = max(22, 18 * max(len(agents_in_slot), 1))
            ws_week.row_dimensions[row_idx].height = row_height
            row_idx += 1

        widths = [16, 18, 14, 18, 42]
        for col_idx, width in enumerate(widths, start=1):
            ws_week.column_dimensions[get_column_letter(col_idx)].width = width


    # Sheet data mentah untuk arsip atau olah ulang.
    # Kolom status request khusus tidak ikut diekspor agar output tetap bersih.
    ws_detail = wb.create_sheet(title="Data Detail")
    detail_export_df = assignments_df.drop(columns=["Urgent"], errors="ignore")
    if detail_export_df.empty:
        ws_detail["A1"] = "Belum ada data."
    else:
        for col_idx, col_name in enumerate(detail_export_df.columns, start=1):
            cell = ws_detail.cell(row=1, column=col_idx, value=col_name)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor=gray)
            cell.border = border
        for row_idx, row in enumerate(detail_export_df.itertuples(index=False), start=2):
            for col_idx, value in enumerate(row, start=1):
                cell = ws_detail.cell(row=row_idx, column=col_idx, value=value)
                cell.border = border
        for col_idx in range(1, len(detail_export_df.columns) + 1):
            ws_detail.column_dimensions[get_column_letter(col_idx)].width = 18


    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# -----------------------------------------------------------------------------
# OUTPUT GAMBAR PNG DAN PDF
# -----------------------------------------------------------------------------

_FONT_PATH_CACHE: Dict[bool, Optional[str]] = {True: None, False: None}


def _try_font(path_or_name: str, size: int) -> Optional[ImageFont.FreeTypeFont]:
    """Mencoba membuka font TrueType/OpenType dengan aman."""
    try:
        return ImageFont.truetype(path_or_name, size=size)
    except Exception:
        return None


def _discover_system_font(bold: bool) -> Optional[str]:
    """Mencari font yang tersedia pada Windows, Linux, dan Streamlit Cloud.

    Masalah yang sering terjadi di server deploy adalah path font berbeda dengan
    laptop lokal. Jika Pillow memakai font default bawaan, teks menjadi sangat
    kecil dan poster terlihat rusak. Fungsi ini melakukan pencarian font sistem
    secara lebih luas agar output tetap konsisten di Streamlit Community Cloud.
    """
    if _FONT_PATH_CACHE.get(bold):
        return _FONT_PATH_CACHE[bold]

    preferred_keywords = (
        [
            "dejavusans-bold", "liberationsans-bold", "arialbd", "arial-bold",
            "notosans-bold", "freesansbold", "roboto-bold", "opensans-bold",
        ]
        if bold
        else [
            "dejavusans", "liberationsans-regular", "arial", "notosans-regular",
            "freesans", "roboto-regular", "opensans-regular",
        ]
    )

    search_roots = [
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        "/app/.fonts",
        os.path.expanduser("~/.fonts"),
        os.path.expanduser("~/.local/share/fonts"),
        "C:/Windows/Fonts",
    ]

    discovered_files: List[str] = []
    for root in search_roots:
        if not root or not os.path.isdir(root):
            continue
        for current_root, _, files in os.walk(root):
            for filename in files:
                if filename.lower().endswith((".ttf", ".otf")):
                    discovered_files.append(os.path.join(current_root, filename))

    # Prioritaskan font yang paling mirip Arial/DejaVu/Liberation agar tampilan
    # tetap tebal, bersih, dan dekat dengan desain Brighton.
    for keyword in preferred_keywords:
        for font_path in discovered_files:
            compact_name = os.path.basename(font_path).lower().replace(" ", "").replace("_", "")
            if keyword.replace(" ", "").replace("_", "") in compact_name:
                _FONT_PATH_CACHE[bold] = font_path
                return font_path

    # Fallback terakhir: ambil font sans/regular yang tersedia, bukan bitmap
    # default Pillow, supaya ukuran teks tetap mengikuti parameter size.
    fallback_keywords = ["sans", "arial", "liberation", "dejavu", "noto", "free", "roboto"]
    for keyword in fallback_keywords:
        for font_path in discovered_files:
            compact_name = os.path.basename(font_path).lower()
            if keyword in compact_name:
                _FONT_PATH_CACHE[bold] = font_path
                return font_path

    return None


def find_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    """Mencari font sistem yang umum tersedia dengan fallback aman untuk deploy.

    Font tidak disertakan ke paket aplikasi. Aplikasi memakai font sistem yang
    tersedia pada komputer/server. Pencarian dibuat luas agar di Streamlit Cloud
    teks tidak jatuh ke font bitmap kecil bawaan Pillow.
    """
    size = max(8, int(round(size)))

    direct_candidates = (
        [
            "DejaVuSans-Bold.ttf",
            "LiberationSans-Bold.ttf",
            "Arial Bold.ttf",
            "Arial-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
            "C:/Windows/Fonts/arialbd.ttf",
        ]
        if bold
        else [
            "DejaVuSans.ttf",
            "LiberationSans-Regular.ttf",
            "Arial.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
            "C:/Windows/Fonts/arial.ttf",
        ]
    )

    for candidate in direct_candidates:
        font = _try_font(candidate, size)
        if font is not None:
            return font

    discovered_path = _discover_system_font(bold=bold)
    if discovered_path:
        font = _try_font(discovered_path, size)
        if font is not None:
            return font

    # Jika server benar-benar tidak memiliki font TTF/OTF, gunakan default agar
    # aplikasi tidak crash. Namun kondisi ini jarang terjadi setelah pencarian luas.
    return ImageFont.load_default()


def text_size(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont) -> Tuple[int, int]:
    """Mengukur ukuran teks dengan aman untuk berbagai versi Pillow."""
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def draw_brighton_logo(
    draw: ImageDraw.ImageDraw,
    center_x: int,
    y: int,
    logo_text: str,
    font: ImageFont.ImageFont,
    black: str,
    white: str,
) -> Tuple[int, int]:
    """Menggambar logo teks Brighton dengan huruf O berwarna putih.

    Jika teks logo bukan "Brighton", aplikasi tetap menggambar teks custom
    secara normal agar fleksibel untuk hub lain atau kebutuhan internal.
    """
    logo = logo_text.strip() or DEFAULT_LOGO_TEXT
    if logo.casefold() == "brighton":
        parts = [("Bright", black), ("o", white), ("n", black)]
        sizes = [text_size(draw, part, font) for part, _ in parts]
        total_w = sum(w for w, _ in sizes)
        max_h = max(h for _, h in sizes)
        x = center_x - total_w // 2
        for (part, color), (part_w, _) in zip(parts, sizes):
            draw.text((x, y), part, font=font, fill=color)
            x += part_w
        return total_w, max_h

    logo_w, logo_h = text_size(draw, logo, font)
    draw.text((center_x - logo_w // 2, y), logo, font=font, fill=black)
    return logo_w, logo_h


def draw_text_fit(
    draw: ImageDraw.ImageDraw,
    xy: Tuple[int, int],
    text: str,
    font: ImageFont.ImageFont,
    fill: str,
    max_width: int,
    min_size: int = 14,
    bold: bool = True,
) -> None:
    """Menggambar teks dan mengecilkan font otomatis bila teks terlalu panjang."""
    x, y = xy
    active_font = font
    size = getattr(font, "size", min_size)
    while size > min_size and text_size(draw, text, active_font)[0] > max_width:
        size -= 1
        active_font = find_font(size, bold=bold)
    draw.text((x, y), text, font=active_font, fill=fill)


def draw_clock_icon(draw: ImageDraw.ImageDraw, center: Tuple[int, int], radius: int, color: str) -> None:
    """Menggambar ikon jam sederhana agar tidak bergantung pada emoji/font khusus."""
    cx, cy = center
    draw.ellipse((cx - radius, cy - radius, cx + radius, cy + radius), outline=color, width=3)
    draw.line((cx, cy, cx, cy - radius + 5), fill=color, width=3)
    draw.line((cx, cy, cx + radius - 5, cy + 3), fill=color, width=3)


def polygon_day_badge(x: int, y: int, w: int, h: int) -> List[Tuple[int, int]]:
    """Bentuk label hari dengan potongan miring seperti referensi Brighton."""
    return [(x, y), (x + w, y), (x + w - 14, y + h), (x, y + h)]


def get_canvas_size(orientation: str) -> Tuple[int, int]:
    """Ukuran kanvas poster.

    Portrait dibuat mendekati rasio A4 vertikal, landscape mendekati A4 horizontal.
    Ukuran pixel besar agar hasil PNG/PDF tetap tajam saat dicetak.
    """
    if orientation == "Landscape":
        return 1680, 1188
    return 1080, 1528


def collect_day_data(
    slots: List[ShiftSlot],
    schedule_for_week: Dict[str, List[Tuple[str, bool]]],
) -> List[Tuple[date, List[ShiftSlot], int]]:
    """Mengelompokkan slot per hari dan menghitung jumlah baris agen maksimum."""
    date_to_slots: Dict[date, List[ShiftSlot]] = {}
    for slot in slots:
        date_to_slots.setdefault(slot.tanggal, []).append(slot)

    day_data: List[Tuple[date, List[ShiftSlot], int]] = []
    for d in sorted(date_to_slots):
        day_slots = sorted(date_to_slots[d], key=lambda s: s.shift_index)
        max_agents = max((len(schedule_for_week.get(slot.key, [])) for slot in day_slots), default=1)
        day_data.append((d, day_slots, max(max_agents, 1)))
    return day_data


def compute_image_layout(
    orientation: str,
    slots: List[ShiftSlot],
    schedule_for_week: Dict[str, List[Tuple[str, bool]]],
) -> Dict[str, int]:
    """Menghitung layout responsif agar poster tetap rapi walau agen banyak.

    Perbaikan penting:
    - Landscape memakai mode compact agar 6 hari kerja dengan ±37 agen tetap muat.
    - Font, ikon, garis, dan seluruh koordinat tetap diskalakan bersama agar PNG/PDF HD.
    - Area catatan selalu disediakan dari awal sehingga tidak menimpa jadwal Sabtu.
    """
    scale = EXPORT_RENDER_SCALE
    base_width, base_height = get_canvas_size(orientation)
    width, height = base_width * scale, base_height * scale
    is_landscape = orientation == "Landscape"

    def px(value: int) -> int:
        return int(round(value * scale))

    if is_landscape:
        # Mode compact khusus landscape. Tanpa ini, minggu penuh Senin-Sabtu
        # dengan agen 35+ akan terlihat terlalu renggang atau catatan menimpa Sabtu.
        margin_x = px(44)
        brand_h = px(72)
        title_h = px(122)
        notes_h = px(70)
        footer_h = px(16)
        body_top = brand_h + title_h + px(12)
        note_badge_reserved_h = px(34)
        header_h = px(42)
        row_h = px(25)
        gap_h = px(7)
        agent_font_size = px(17)
        time_font_size = px(14)
        col_gap = px(34)
        time_w = px(170)
        day_badge_w = px(190)
        day_badge_h = px(40)
    else:
        margin_x = px(54)
        brand_h = px(100)
        title_h = px(128)
        notes_h = px(118)
        footer_h = px(18)
        body_top = brand_h + title_h + px(18)
        note_badge_reserved_h = px(50)
        header_h = px(58)
        row_h = px(32)
        gap_h = px(16)
        agent_font_size = px(24)
        time_font_size = px(22)
        col_gap = px(28)
        time_w = px(202)
        day_badge_w = px(192)
        day_badge_h = px(56)

    body_bottom = height - notes_h - footer_h - px(24) - note_badge_reserved_h
    available_body_h = max(px(420), body_bottom - body_top)

    day_data = collect_day_data(slots, schedule_for_week)

    def required_height(candidate_row_h: int, candidate_gap_h: int) -> int:
        total = 0
        for _, _, max_agents in day_data:
            total += header_h + px(8) + max(candidate_row_h * max_agents + px(8), candidate_row_h + px(8)) + candidate_gap_h
        return total

    min_row_h = px(19 if is_landscape else 22)
    min_agent_font = px(13 if is_landscape else 15)
    min_time_font = px(12 if is_landscape else 14)
    while required_height(row_h, gap_h) > available_body_h and row_h > min_row_h:
        row_h -= scale
        agent_font_size = max(min_agent_font, agent_font_size - scale)
        time_font_size = max(min_time_font, time_font_size - scale)

    min_gap_h = px(4 if is_landscape else 6)
    while required_height(row_h, gap_h) > available_body_h and gap_h > min_gap_h:
        gap_h -= scale

    # Cadangan terakhir: landscape yang sangat penuh tetap dibuat aman dengan
    # mengurangi tinggi notes, bukan menimpa jadwal.
    while required_height(row_h, gap_h) > available_body_h and notes_h > px(52):
        notes_h -= px(4)
        body_bottom = height - notes_h - footer_h - px(20) - note_badge_reserved_h
        available_body_h = max(px(420), body_bottom - body_top)

    shift_area_x = margin_x
    shift_area_w = width - 2 * margin_x
    col_w = (shift_area_w - col_gap) // 2

    return {
        "width": width,
        "height": height,
        "margin_x": margin_x,
        "brand_h": brand_h,
        "title_h": title_h,
        "notes_h": notes_h,
        "footer_h": footer_h,
        "body_top": body_top,
        "body_bottom": body_bottom,
        "header_h": header_h,
        "row_h": row_h,
        "gap_h": gap_h,
        "agent_font_size": agent_font_size,
        "time_font_size": time_font_size,
        "col_gap": col_gap,
        "shift_area_x": shift_area_x,
        "shift_area_w": shift_area_w,
        "col_w": col_w,
        "time_w": time_w,
        "day_badge_w": day_badge_w,
        "day_badge_h": day_badge_h,
        "scale": scale,
    }


def draw_shift_column(
    draw: ImageDraw.ImageDraw,
    x: int,
    y: int,
    col_w: int,
    time_w: int,
    row_h: int,
    time_label: str,
    agents: List[Tuple[str, bool]],
    time_font: ImageFont.ImageFont,
    agent_font: ImageFont.ImageFont,
    small_font: ImageFont.ImageFont,
    colors: Dict[str, str],
) -> None:
    """Menggambar satu kolom shift dengan alignment presisi.

    Jam, icon jam, bullet pertama, dan nama agen pertama memakai baseline yang sama,
    sehingga tidak ada lagi tampilan jam naik/turun atau bullet menempel pada jam.
    Garis bawah area jam dihilangkan agar tidak tampak terpotong di sekitar ikon.
    """
    line_color = colors["line_gray"]
    yellow = colors["yellow"]
    black = colors["black"]
    gray_text = colors["gray_text"]
    scale = max(1, EXPORT_RENDER_SCALE)

    first_row_y = y
    clock_center = (x + 22 * scale, first_row_y + row_h // 2)
    draw_clock_icon(draw, clock_center, 15 * scale, yellow)
    draw_text_fit(
        draw,
        (x + 50 * scale, first_row_y + 4 * scale),
        time_label,
        time_font,
        gray_text,
        time_w - 54 * scale,
        min_size=13 * scale,
        bold=True,
    )

    list_x = x + time_w + 28 * scale
    list_w = col_w - time_w - 28 * scale
    if not agents:
        agents = [("", False)]

    for idx, (agent, _is_urgent) in enumerate(agents):
        row_y = y + idx * row_h
        baseline = row_y + row_h - 2 * scale
        bullet_cy = row_y + row_h // 2
        draw.ellipse((list_x, bullet_cy - 5 * scale, list_x + 10 * scale, bullet_cy + 5 * scale), fill=black)
        draw_text_fit(
            draw,
            (list_x + 30 * scale, row_y + 2 * scale),
            agent,
            agent_font,
            black,
            max_width=max(60 * scale, list_w - 30 * scale),
            min_size=12 * scale,
            bold=True,
        )
        # Status urgent tidak diberi badge pada jadwal agar tampilan tetap bersih untuk agen.
        draw.line((list_x - 6 * scale, baseline, list_x + list_w, baseline), fill=line_color, width=max(2, 3 * scale // 2))


def render_week_image(
    week_no: int,
    week_dates: List[date],
    slots: List[ShiftSlot],
    schedule_for_week: Dict[str, List[Tuple[str, bool]]],
    hub_name: str,
    logo_text: str,
    notes: List[str],
    orientation: str,
) -> Image.Image:
    """Membuat poster jadwal satu minggu dalam bentuk gambar PNG.

    Fungsi ini sengaja tidak bergantung pada HTML/CSS agar output stabil di semua
    browser, saat preview, saat download gambar, dan saat dibuat PDF.
    """
    layout = compute_image_layout(orientation, slots, schedule_for_week)
    width = layout["width"]
    height = layout["height"]
    margin_x = layout["margin_x"]
    scale = layout.get("scale", 1)

    colors = {
        "yellow": "#FFD10A",
        "black": "#111111",
        "white": "#FFFFFF",
        "soft_gray": "#F3F3F3",
        "line_gray": "#D7D7D7",
        "muted_line": "#F1DC75",
        "gray_text": "#444444",
    }

    img = Image.new("RGB", (width, height), colors["white"])
    draw = ImageDraw.Draw(img)

    brand_font = find_font((72 if orientation == "Landscape" else 64) * scale, bold=True)
    hub_font = find_font((28 if orientation == "Landscape" else 22) * scale, bold=True)
    title_font = find_font((48 if orientation == "Landscape" else 43) * scale, bold=True)
    range_font = find_font((32 if orientation == "Landscape" else 27) * scale, bold=True)
    week_font = find_font((22 if orientation == "Landscape" else 18) * scale, bold=True)
    day_font = find_font((34 if orientation == "Landscape" else 31) * scale, bold=True)
    date_font = find_font((31 if orientation == "Landscape" else 28) * scale, bold=True)
    special_font = find_font((21 if orientation == "Landscape" else 16) * scale, bold=True)
    time_font = find_font(layout["time_font_size"], bold=True)
    agent_font = find_font(layout["agent_font_size"], bold=True)
    small_font = find_font(13 * scale, bold=True)
    note_title_font = find_font(24 * scale, bold=True)
    note_font = find_font((19 if orientation == "Landscape" else 18) * scale, bold=False)
    note_bold_font = find_font((19 if orientation == "Landscape" else 18) * scale, bold=True)

    # Brand bar
    brand_h = layout["brand_h"]
    draw.rectangle((0, 0, width, brand_h), fill=colors["yellow"])
    logo = logo_text.strip() or DEFAULT_LOGO_TEXT
    # Ukur tinggi logo terlebih dahulu agar posisi vertikal tetap presisi.
    if logo.casefold() == "brighton":
        logo_h = max(text_size(draw, part, brand_font)[1] for part in ["Bright", "o", "n"])
    else:
        _, logo_h = text_size(draw, logo, brand_font)
    logo_y = (brand_h - logo_h) // 2 - 4 * scale
    draw_brighton_logo(
        draw=draw,
        center_x=width // 2,
        y=logo_y,
        logo_text=logo,
        font=brand_font,
        black=colors["black"],
        white=colors["white"],
    )
    hub_text = hub_name.strip() or DEFAULT_HUB_NAME
    hub_w, hub_h = text_size(draw, hub_text, hub_font)
    draw.text((width - margin_x - hub_w, brand_h - hub_h - 12 * scale), hub_text, font=hub_font, fill=colors["white"])

    # Judul
    title_y = brand_h + 24 * scale
    title = "FLOOR TIME SCHEDULE"
    title_w, title_h = text_size(draw, title, title_font)
    draw.text(((width - title_w) // 2, title_y), title, font=title_font, fill=colors["black"])
    line_w = int(width * (0.56 if orientation == "Landscape" else 0.58))
    line_y = title_y + title_h + 13 * scale
    draw.rectangle(((width - line_w) // 2, line_y, (width + line_w) // 2, line_y + 3 * scale), fill=colors["muted_line"])
    range_text = week_range_label(week_dates)
    range_w, range_h = text_size(draw, range_text, range_font)
    draw.text(((width - range_w) // 2, line_y + 10 * scale), range_text, font=range_font, fill=colors["black"])
    week_text = f"MINGGU {week_no}"
    week_w, week_h = text_size(draw, week_text, week_font)
    draw.text(((width - week_w) // 2, line_y + 10 * scale + range_h + 12 * scale), week_text, font=week_font, fill=colors["gray_text"])

    # Isi jadwal per hari
    y = layout["body_top"]
    day_data = collect_day_data(slots, schedule_for_week)
    for d, day_slots, max_agents in day_data:
        badge_w = layout["day_badge_w"]
        badge_h = layout["day_badge_h"]
        header_y = y
        line_y = header_y + badge_h - 4 * scale
        draw.rectangle((margin_x, line_y, width - margin_x, line_y + max(4, 5 * scale // 2)), fill=colors["yellow"])
        draw.polygon(polygon_day_badge(margin_x, header_y, badge_w, badge_h), fill=colors["yellow"])
        day_name = DAY_NAMES_ID[d.weekday()]
        day_w, day_h = text_size(draw, day_name, day_font)
        draw.text((margin_x + (badge_w - day_w) // 2 - 4 * scale, header_y + (badge_h - day_h) // 2 - 5 * scale), day_name, font=day_font, fill=colors["black"])
        date_label = format_date_id(d)
        draw.text((margin_x + badge_w + 18 * scale, header_y + 12 * scale), date_label, font=date_font, fill=colors["black"])

        if d.weekday() == 5:
            special_text = "Jam Operasional Khusus"
            sp_w = (300 if orientation == "Landscape" else 230) * scale
            sp_h = (38 if orientation == "Landscape" else 32) * scale
            sp_x = width - margin_x - sp_w
            sp_y = header_y + 10 * scale
            draw.rectangle((sp_x, sp_y, sp_x + sp_w, sp_y + sp_h), fill=colors["yellow"])
            tw, th = text_size(draw, special_text, special_font)
            draw.text((sp_x + (sp_w - tw) // 2, sp_y + (sp_h - th) // 2 - 2 * scale), special_text, font=special_font, fill=colors["white"])

        content_y = header_y + badge_h + 12 * scale
        col1_x = layout["shift_area_x"]
        col2_x = layout["shift_area_x"] + layout["col_w"] + layout["col_gap"]
        col_w = layout["col_w"]
        time_w = layout["time_w"]
        row_h = layout["row_h"]

        for idx, slot in enumerate(day_slots[:2]):
            col_x = col1_x if idx == 0 else col2_x
            draw_shift_column(
                draw=draw,
                x=col_x,
                y=content_y,
                col_w=col_w,
                time_w=time_w,
                row_h=row_h,
                time_label=slot.time_label,
                agents=schedule_for_week.get(slot.key, []),
                time_font=time_font,
                agent_font=agent_font,
                small_font=small_font,
                colors=colors,
            )

        shift_h = max_agents * row_h + 8 * scale
        y = content_y + shift_h + layout["gap_h"]

    if not day_data:
        empty_font = find_font(22 * scale, bold=True)
        draw.text((margin_x, y + 30 * scale), "Tidak ada jadwal aktif pada minggu ini.", font=empty_font, fill=colors["gray_text"])

    # Catatan
    notes_h = layout["notes_h"]
    footer_h = layout["footer_h"]
    note_box_x = margin_x + 30 * scale
    note_box_w = width - (margin_x + 30 * scale) * 2
    note_box_y = height - footer_h - notes_h - 20 * scale
    draw.rectangle((note_box_x, note_box_y, note_box_x + note_box_w, note_box_y + notes_h), fill=colors["soft_gray"])
    badge_w = 220 * scale
    badge_h = 44 * scale
    draw.polygon(polygon_day_badge(note_box_x, note_box_y - badge_h + 4, badge_w, badge_h), fill=colors["yellow"])
    draw.text((note_box_x + 16 * scale, note_box_y - badge_h + 13 * scale), "CATATAN:", font=note_title_font, fill=colors["black"])

    clean_notes = [note.strip() for note in notes if note.strip()] or ["Agen yang mendapatkan jadwal floor time masih berada di kantor."]
    text_y = note_box_y + 24 * scale
    for note in clean_notes[:3]:
        bullet_x = note_box_x + 20 * scale
        draw.ellipse((bullet_x, text_y + 8 * scale, bullet_x + 8 * scale, text_y + 16 * scale), fill=colors["black"])
        # Penekanan sederhana untuk frasa operasional penting.
        normal_text = note
        draw_text_fit(draw, (bullet_x + 20 * scale, text_y), normal_text, note_font, colors["black"], note_box_w - 52 * scale, min_size=13 * scale, bold=False)
        text_y += 32 * scale

    # Footer kuning bersih untuk output jadwal.
    # Credit pembuat aplikasi ditampilkan hanya di halaman aplikasi, bukan di file output.
    draw.rectangle((0, height - footer_h, width, height), fill=colors["yellow"])
    return img


def render_all_week_images(
    weeks: Dict[int, List[date]],
    slots_by_week: Dict[int, List[ShiftSlot]],
    schedule: Dict[int, Dict[str, List[Tuple[str, bool]]]],
    hub_name: str,
    logo_text: str,
    notes: List[str],
    orientation: str,
) -> Dict[int, Image.Image]:
    """Render seluruh minggu menjadi dictionary nomor minggu -> gambar."""
    images: Dict[int, Image.Image] = {}
    for week_no, week_dates in weeks.items():
        images[week_no] = render_week_image(
            week_no=week_no,
            week_dates=week_dates,
            slots=slots_by_week.get(week_no, []),
            schedule_for_week=schedule.get(week_no, {}),
            hub_name=hub_name,
            logo_text=logo_text,
            notes=notes,
            orientation=orientation,
        )
    return images


def build_images_zip(week_images: Dict[int, Image.Image], year: int, month: int) -> bytes:
    """Membuat ZIP berisi gambar PNG per minggu."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for week_no, img in week_images.items():
            img_buffer = io.BytesIO()
            img.save(img_buffer, format="PNG", dpi=(300, 300), compress_level=1)
            zf.writestr(f"floor_time_{year}-{month:02d}_minggu_{week_no}.png", img_buffer.getvalue())
    return buffer.getvalue()


def build_pdf_from_images(week_images: Dict[int, Image.Image]) -> bytes:
    """Membuat PDF multi-page dari gambar poster mingguan."""
    if not week_images:
        return b""
    ordered_images = [img.convert("RGB") for _, img in sorted(week_images.items())]
    buffer = io.BytesIO()
    first, rest = ordered_images[0], ordered_images[1:]
    first.save(buffer, format="PDF", save_all=True, append_images=rest, resolution=300.0)
    return buffer.getvalue()


# -----------------------------------------------------------------------------
# UI STREAMLIT
# -----------------------------------------------------------------------------

def inject_app_style() -> None:
    """Style kecil untuk tampilan Streamlit."""
    st.markdown(
        """
        <style>
            .main-title {
                padding: 18px 22px;
                background: #ffd10a;
                border-radius: 18px;
                color: #111111;
                margin-bottom: 18px;
            }
            .main-title h1 {
                margin: 0;
                font-size: 34px;
                line-height: 1.1;
            }
            .main-title p {
                margin: 8px 0 0 0;
                font-size: 15px;
            }
            .app-credit {
                text-align: center;
                margin: 10px 0 4px 0;
                padding: 0;
                font-size: 13px;
                font-weight: 600;
                color: #666666;
                letter-spacing: .2px;
                line-height: 1.2;
            }
            .block-container {
                padding-bottom: 0.6rem !important;
            }
            .brand-o-white {
                color: #ffffff;
            }
            .metric-card {
                border: 1px solid #eeeeee;
                padding: 14px 16px;
                border-radius: 14px;
                background: #ffffff;
            }
            .small-muted {
                color: #666666;
                font-size: 13px;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )



def render_app_credit() -> None:
    """Menampilkan credit kecil di aplikasi.

    Credit hanya tampil di aplikasi, tidak ikut masuk ke output Excel, PNG, atau PDF.
    Jaraknya dibuat rapat dan natural seperti footer aplikasi profesional.
    """
    st.markdown('<div class="app-credit">Created by rh</div>', unsafe_allow_html=True)


def initialize_session_state() -> None:
    """Menyiapkan state agar hasil generate tidak hilang saat UI berubah."""
    defaults = {
        "assignments_df": pd.DataFrame(),
        "schedule": {},
        "warnings": [],
        "validation_messages": [],
        "validation_ok": False,
        "html_document": "",
        "excel_bytes": b"",
        "image_zip_bytes": b"",
        "pdf_bytes": b"",
        "week_images": {},
        "output_orientation": "Portrait",
        "slots_by_week": {},
        "weeks": {},
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def main() -> None:
    """Entry point aplikasi Streamlit."""
    st.set_page_config(page_title=APP_TITLE, page_icon="🟨", layout="wide")
    inject_app_style()
    initialize_session_state()

    st.markdown(
        """
        <div class="main-title">
            <h1>Bright<span class="brand-o-white">o</span>n Floor Time Schedule Generator</h1>
            <p>Generator jadwal floor time bulanan dengan tanggal merah, request urgent, randomisasi penuh, validasi anti-double, dan output siap cetak.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    today = datetime.today().date()

    with st.sidebar:
        st.header("⚙️ Pengaturan Jadwal")
        hub_name = st.text_input("Nama Hub", value=DEFAULT_HUB_NAME)
        logo_text = st.text_input("Teks Logo", value=DEFAULT_LOGO_TEXT)

        selected_year = st.number_input("Tahun", min_value=2024, max_value=2100, value=today.year, step=1)
        selected_month = st.selectbox(
            "Bulan",
            options=list(MONTH_NAMES_ID.keys()),
            index=today.month - 1,
            format_func=lambda m: MONTH_NAMES_ID[m],
        )

        st.divider()
        st.subheader("Kapasitas Shift")
        weekday_capacity = st.number_input("Kapasitas awal per shift Senin-Jumat", min_value=1, max_value=30, value=4, step=1)
        saturday_capacity = st.number_input("Kapasitas awal per shift Sabtu", min_value=1, max_value=30, value=4, step=1)
        auto_expand_capacity = st.checkbox("Naikkan kapasitas otomatis bila agen lebih banyak", value=True)

        st.divider()
        st.subheader("Orientasi Output")
        output_orientation = st.radio(
            "Pilih bentuk hasil poster",
            options=["Portrait", "Landscape"],
            index=0,
            horizontal=True,
            help="Portrait cocok untuk format story/lembar vertikal. Landscape cocok untuk jadwal yang ingin terlihat lebih lebar.",
        )

        st.divider()
        st.subheader("Randomisasi")
        use_seed = st.checkbox("Gunakan kode audit agar hasil bisa diulang", value=False)
        seed_text = ""
        if use_seed:
            seed_text = st.text_input("Kode audit random", value=f"{selected_year}-{selected_month:02d}-floor-time")
        else:
            st.caption("Tanpa kode audit, hasil akan lebih acak setiap tombol Generate ditekan.")

    work_dates = get_month_work_dates(int(selected_year), int(selected_month))
    weeks = group_dates_by_calendar_week(work_dates)
    date_labels = [date_option_label(d) for d in work_dates]
    label_to_date = {date_option_label(d): d for d in work_dates}

    tab_input, tab_result, tab_export, tab_help = st.tabs(["1. Input Data", "2. Hasil Jadwal", "3. Download", "4. Panduan"])

    with tab_input:
        left_col, right_col = st.columns([1.08, 0.92], gap="large")

        with left_col:
            st.subheader("Daftar Agen")
            st.caption("Masukkan satu agen per baris. Format dengan kode seperti PAULA (DUCC) tetap didukung.")
            sample_text = """PAULA (DUCC)
AZWAR (IBEX)
MONIKA (NDEP)
WINARDI (ECFE)
KATRIN (VHEL)
SOVIE (KNIP)
ANNA (NKRL)"""
            raw_agents = st.text_area("Input nama agen", value=sample_text, height=250)
            uploaded_file = st.file_uploader("Atau upload CSV/Excel daftar agen", type=["csv", "xlsx", "xls"])

            text_agents, duplicates = parse_agents_from_text(raw_agents)
            uploaded_agents = read_agents_from_upload(uploaded_file) if uploaded_file else []

            # Bila user upload file, gabungkan dengan textarea dan tetap hilangkan duplikat.
            combined_agents = []
            seen_agents = set()
            for agent in text_agents + uploaded_agents:
                key = agent.casefold()
                if key not in seen_agents:
                    seen_agents.add(key)
                    combined_agents.append(agent)

            st.info(f"Total agen unik: {len(combined_agents)}")
            if duplicates:
                st.warning(f"Duplikat dari textarea diabaikan: {', '.join(duplicates)}")

        with right_col:
            st.subheader("Tanggal Merah atau Kantor Tutup")
            st.caption("Pilih semua tanggal yang tidak boleh mendapat jadwal floor time.")
            selected_closed_labels = st.multiselect("Pilih tanggal libur", options=date_labels)
            closed_dates = [label_to_date[label] for label in selected_closed_labels]

            st.subheader("Request Jadwal Khusus Internal")
            st.caption("Gunakan hanya untuk kondisi penting. Informasi ini hanya untuk admin dan tidak tampil pada jadwal publik.")
            urgent_template = pd.DataFrame(columns=["Agen", "Tanggal", "Shift", "Catatan"])
            urgent_df = st.data_editor(
                urgent_template,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Agen": st.column_config.SelectboxColumn("Agen", options=combined_agents if combined_agents else [""], required=False),
                    "Tanggal": st.column_config.SelectboxColumn("Tanggal", options=date_labels if date_labels else [""], required=False),
                    "Shift": st.column_config.SelectboxColumn("Shift", options=["Pagi", "Siang"], required=False),
                    "Catatan": st.column_config.TextColumn("Catatan", help="Opsional, untuk arsip internal."),
                },
                hide_index=True,
            )

        notes_default = """Agen yang mendapatkan jadwal floor time masih berada di kantor.
Penanganan klien berdasarkan agen yang melakukan absensi pertama."""
        notes_raw = st.text_area("Catatan pada poster jadwal", value=notes_default, height=90)
        notes = [line.strip() for line in notes_raw.splitlines() if line.strip()]

        st.divider()
        generate_clicked = st.button("Generate Jadwal Bulanan", type="primary", use_container_width=True)

        if generate_clicked:
            if not combined_agents:
                st.error("Daftar agen belum diisi.")
                st.stop()

            urgent_requests, urgent_errors = parse_urgent_requests(urgent_df, label_to_date, combined_agents)
            if urgent_errors:
                for err in urgent_errors:
                    st.error(err)
                st.stop()

            slots_by_week = build_slots_by_week(
                weeks=weeks,
                closed_dates=closed_dates,
                weekday_capacity=int(weekday_capacity),
                saturday_capacity=int(saturday_capacity),
            )

            assignments_df, schedule, capacity_by_week, warnings = generate_schedule(
                agents=combined_agents,
                slots_by_week=slots_by_week,
                urgent_requests=urgent_requests,
                auto_expand_capacity=auto_expand_capacity,
                seed_text=seed_text if use_seed else None,
            )

            active_weeks = [week_no for week_no, slots in slots_by_week.items() if slots]
            validation_ok, validation_messages = validate_schedule(assignments_df, combined_agents, active_weeks)

            html_document = build_full_html_document(
                weeks=weeks,
                slots_by_week=slots_by_week,
                schedule=schedule,
                hub_name=hub_name,
                logo_text=logo_text,
                notes=notes,
            )
            excel_bytes = build_excel_file(assignments_df, weeks, slots_by_week, schedule, hub_name)
            week_images = render_all_week_images(
                weeks=weeks,
                slots_by_week=slots_by_week,
                schedule=schedule,
                hub_name=hub_name,
                logo_text=logo_text,
                notes=notes,
                orientation=output_orientation,
            )
            image_zip_bytes = build_images_zip(week_images, int(selected_year), int(selected_month))
            pdf_bytes = build_pdf_from_images(week_images)

            st.session_state["assignments_df"] = assignments_df
            st.session_state["schedule"] = schedule
            st.session_state["warnings"] = warnings
            st.session_state["validation_messages"] = validation_messages
            st.session_state["validation_ok"] = validation_ok
            st.session_state["html_document"] = html_document
            st.session_state["excel_bytes"] = excel_bytes
            st.session_state["week_images"] = week_images
            st.session_state["image_zip_bytes"] = image_zip_bytes
            st.session_state["pdf_bytes"] = pdf_bytes
            st.session_state["output_orientation"] = output_orientation
            st.session_state["slots_by_week"] = slots_by_week
            st.session_state["weeks"] = weeks
            st.session_state["hub_name"] = hub_name
            st.session_state["logo_text"] = logo_text
            st.session_state["notes"] = notes

            if validation_ok:
                st.success("Jadwal berhasil digenerate dan lolos validasi anti-double.")
            else:
                st.error("Jadwal berhasil dibuat, tetapi ada validasi yang perlu dicek.")

        render_app_credit()

    with tab_result:
        st.subheader("Preview Jadwal")
        assignments_df = st.session_state.get("assignments_df", pd.DataFrame())
        schedule = st.session_state.get("schedule", {})
        slots_by_week = st.session_state.get("slots_by_week", {})
        weeks_state = st.session_state.get("weeks", {})

        if assignments_df.empty and not schedule:
            st.info("Belum ada jadwal. Isi data pada tab Input Data lalu klik Generate.")
        else:
            metric_cols = st.columns(4)
            total_agents_in_first_week = assignments_df[assignments_df["Minggu"] == assignments_df["Minggu"].min()]["Agen"].nunique() if not assignments_df.empty else 0
            metric_cols[0].metric("Agen", total_agents_in_first_week)
            metric_cols[1].metric("Minggu", len(weeks_state))
            metric_cols[2].metric("Total Baris Jadwal", len(assignments_df))
            active_slot_count = sum(len(slots) for slots in slots_by_week.values()) if slots_by_week else 0
            metric_cols[3].metric("Slot Aktif", active_slot_count)

            validation_ok = st.session_state.get("validation_ok", False)
            validation_messages = st.session_state.get("validation_messages", [])
            if validation_ok:
                st.success("\n".join(validation_messages))
            else:
                for msg in validation_messages:
                    st.warning(msg)

            for warning in st.session_state.get("warnings", []):
                st.warning(warning)

            if not assignments_df.empty:
                with st.expander("Lihat data detail jadwal"):
                    detail_preview_df = assignments_df.drop(columns=["Urgent"], errors="ignore")
                    st.dataframe(detail_preview_df, use_container_width=True, hide_index=True)

            week_numbers = list(weeks_state.keys())
            if week_numbers:
                selected_week = st.selectbox(
                    "Pilih minggu untuk preview",
                    options=week_numbers,
                    format_func=lambda w: f"Minggu {w} - {week_range_label(weeks_state[w])}",
                )
                week_images = st.session_state.get("week_images", {})
                if selected_week in week_images:
                    st.image(week_images[selected_week], use_container_width=True)
                else:
                    st.info("Preview gambar belum tersedia. Silakan generate ulang jadwal.")

        render_app_credit()

    with tab_export:
        st.subheader("Download Hasil")
        assignments_df = st.session_state.get("assignments_df", pd.DataFrame())
        if assignments_df.empty:
            st.info("Generate jadwal terlebih dahulu agar file download tersedia.")
        else:
            month_file = f"{int(selected_year)}-{int(selected_month):02d}"
            excel_bytes = st.session_state.get("excel_bytes", b"")
            image_zip_bytes = st.session_state.get("image_zip_bytes", b"")
            pdf_bytes = st.session_state.get("pdf_bytes", b"")
            orientation_used = st.session_state.get("output_orientation", "Portrait")

            st.caption(f"Orientasi file yang terakhir digenerate: **{orientation_used}**. Ubah orientasi di sidebar lalu generate ulang bila ingin bentuk lain.")

            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button(
                    label="Download Excel",
                    data=excel_bytes,
                    file_name=f"floor_time_{month_file}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with c2:
                st.download_button(
                    label="Download Gambar PNG",
                    data=image_zip_bytes,
                    file_name=f"floor_time_gambar_{month_file}.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
            with c3:
                st.download_button(
                    label="Download PDF",
                    data=pdf_bytes,
                    file_name=f"floor_time_{month_file}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

            st.caption("File gambar berisi ZIP karena jadwal bulanan terdiri dari beberapa poster mingguan. PDF sudah dibuat multi-page sesuai jumlah minggu.")

        render_app_credit()

    with tab_help:
        st.subheader("Panduan Penggunaan")
        st.markdown(
            """
            **Alur kerja yang disarankan:**

            1. Pilih bulan dan tahun jadwal.
            2. Masukkan daftar agen yang berhak mendapatkan jadwal floor time.
            3. Pilih tanggal merah atau tanggal kantor tutup.
            4. Masukkan request jadwal khusus bila ada. Gunakan seperlunya agar hasil tetap adil.
            5. Klik **Generate Jadwal Bulanan**.
            6. Cek tab **Hasil Jadwal**. Pastikan validasi anti-double berhasil.
            7. Download Excel untuk arsip, PNG untuk gambar, dan PDF untuk cetak.

            **Aturan yang diterapkan aplikasi:**

            - Setiap agen hanya mendapat 1 jadwal dalam 1 minggu aktif.
            - Tanggal merah tidak akan diberikan jadwal.
            - Request jadwal khusus diprioritaskan secara internal, tetapi tidak diberi label pada jadwal publik.
            - Bila jumlah agen lebih banyak dari kapasitas shift, aplikasi dapat menaikkan kapasitas otomatis.
            - Minggu terakhir otomatis diteruskan sampai Sabtu walaupun tanggalnya masuk bulan berikutnya.

            **Catatan penting:**

            Bila ingin hasil random yang bisa diulang untuk audit, aktifkan **Gunakan kode audit** di sidebar. Bila tidak aktif, hasil random akan berubah setiap kali tombol Generate ditekan.
            """
        )

        render_app_credit()


if __name__ == "__main__":
    main()
