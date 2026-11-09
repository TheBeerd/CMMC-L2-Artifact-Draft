#!/usr/bin/env python3
"""
CMMC L2 SSP Drafting Tool - Local Application
Generates System Security Plans for CMMC Level 2 compliance
"""

import os
import re
import json
import uuid
import shutil
import secrets
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, session
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.data_source import AxDataSource, StrRef
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

from gap_seed import GAP_QUESTIONS
from sra_seed import SRA_FAMILIES, SRA_SHEETS, SRA_SHEET_LABELS, SRA_ITEMS
from controls_narratives_seed import CONTROLS_NARRATIVES

app = Flask(__name__)
# No feature in this app actually stores data in flask.session — nothing
# depends on the key surviving a restart, so a fresh random key per process
# is simpler and safer than a static one baked into source. Set SSP_SECRET_KEY
# in the environment if a stable key is ever needed (e.g. multi-worker hosting).
app.secret_key = os.environ.get('SSP_SECRET_KEY') or secrets.token_hex(32)

# Storage directory (local only)
DATA_DIR = Path(__file__).parent / 'data'
DATA_DIR.mkdir(exist_ok=True)
UPLOADS_DIR = Path(__file__).parent / 'uploads'
UPLOADS_DIR.mkdir(exist_ok=True)

SSP_FILE = DATA_DIR / 'ssp_data.json'
ASSET_FILE = DATA_DIR / 'asset_data.json'
POAM_FILE = DATA_DIR / 'poam_data.json'
SRA_FILE = DATA_DIR / 'sra_data.json'

# ─────────────────────────────────────────────────────────────
# CMMC L2 Asset Categories (per CMMC Scoping Guide L2 v2.13)
# ─────────────────────────────────────────────────────────────
ASSET_CATEGORIES = {
    "CUI": "CUI Asset – Processes, stores, or transmits CUI",
    "SPA": "Security Protection Asset – Provides security functions/capabilities",
    "CRMA": "Contractor Risk Managed Asset – Can process CUI but controlled by policy",
    "SA_GFE": "Specialized Asset – Government Furnished Equipment (GFE)",
    "SA_IoT": "Specialized Asset – IoT / IIoT Device",
    "SA_OT": "Specialized Asset – Operational Technology (OT)",
    "SA_RI": "Specialized Asset – Restricted Information System",
    "SA_TE": "Specialized Asset – Test Equipment",
    "OOS": "Out-of-Scope Asset – Cannot process CUI; physically/logically separated",
}

PHYSICAL_DEVICE_TYPES = [
    "Badge Reader / PACS Device",
    "Biometric Reader",
    "Security Camera / CCTV",
    "NVR / DVR (Camera Recorder)",
    "Alarm Panel / Controller",
    "Motion Sensor",
    "Door Lock / Electronic Strike",
    "Intercom / Video Doorbell",
    "Guard Station Equipment",
    "Visitor Management Kiosk",
    "Physical Key Management Cabinet",
    "Other Physical Security Device",
]

DEVICE_TYPES = [
    "Workstation (Desktop)",
    "Laptop",
    "Server (Physical)",
    "Server (Virtual/VM)",
    "Network Firewall",
    "Network Switch",
    "Wireless Access Point",
    "Router",
    "VPN Appliance",
    "NAS / Storage Array",
    "Mobile Device (Smartphone)",
    "Mobile Device (Tablet)",
    "USB/Portable Storage",
    "Printer / MFD",
    "Cloud Service (SaaS/IaaS/PaaS)",
    "Domain Controller",
    "Backup Appliance / Device",
] + PHYSICAL_DEVICE_TYPES

PERSONNEL_CATEGORIES = [
    "Full-Time Employee (FTE)",
    "Part-Time Employee",
    "Intern",
    "Contractor",
    "Subcontractor",
    "Temporary / Seasonal",
    "Third-Party Vendor",
    "Consultant",
]

RISK_LEVELS = [
    "1 \u2013 Low",
    "2 \u2013 Low-Moderate",
    "3 \u2013 Moderate",
    "4 \u2013 High",
    "5 \u2013 Critical",
]

# Curated mapping of CMMC control objectives that ask to "identify" or "define"
# a list of users/devices/processes/ports/apps to the corresponding Asset
# Inventory tab. Key: "CONTROL_ID|letter" -> (tab_anchor, label)
# tab_anchor must match a showTab() panel id in assets.html.
# ─────────────────────────────────────────────────────────────
# SPRS / CMMC Level 2 Scoring — per 32 CFR 170.24(c)(2)
# Point values tie to NIST SP 800-171 R2 basic vs derived
# security requirements (FIPS 200 basics; SP 800-53 deriveds).
# ─────────────────────────────────────────────────────────────
SPRS_FIVE_BASIC = {
    'AC.L2-3.1.1','AC.L2-3.1.2','AT.L2-3.2.1','AT.L2-3.2.2','AU.L2-3.3.1',
    'CM.L2-3.4.1','CM.L2-3.4.2','IA.L2-3.5.1','IA.L2-3.5.2','IR.L2-3.6.1',
    'IR.L2-3.6.2','MA.L2-3.7.2','MP.L2-3.8.3','PS.L2-3.9.2','PE.L2-3.10.1',
    'PE.L2-3.10.2','CA.L2-3.12.1','CA.L2-3.12.3','SC.L2-3.13.1','SC.L2-3.13.2',
    'SI.L2-3.14.1','SI.L2-3.14.2','SI.L2-3.14.3',
}
SPRS_FIVE_DERIVED = {
    'AC.L2-3.1.12','AC.L2-3.1.13','AC.L2-3.1.16','AC.L2-3.1.17','AC.L2-3.1.18',
    'AU.L2-3.3.5','CM.L2-3.4.5','CM.L2-3.4.6','CM.L2-3.4.7','CM.L2-3.4.8',
    'IA.L2-3.5.10','MA.L2-3.7.5','MP.L2-3.8.7','RA.L2-3.11.2','SC.L2-3.13.5',
    'SC.L2-3.13.6','SC.L2-3.13.15','SI.L2-3.14.4','SI.L2-3.14.6',
}
SPRS_THREE_BASIC = {
    'AU.L2-3.3.2','MA.L2-3.7.1','MP.L2-3.8.1','MP.L2-3.8.2','PS.L2-3.9.1',
    'RA.L2-3.11.1','CA.L2-3.12.2',
}
SPRS_THREE_DERIVED = {
    'AC.L2-3.1.5','AC.L2-3.1.19','MA.L2-3.7.4','MP.L2-3.8.8','SC.L2-3.13.8',
    'SI.L2-3.14.5','SI.L2-3.14.7',
}
# Partial-credit requirements: deduction is 3 (partial) or 5 (none) per 170.24(c)(2)(i)(A)(4)
SPRS_SPECIAL = {'IA.L2-3.5.3', 'SC.L2-3.13.11'}
# The SSP requirement itself: no point value — without an SSP the assessment
# "could not be completed" per 170.24(c)(2)(i)(A)(5)
SPRS_SSP_REQ = 'CA.L2-3.12.4'

def sprs_point_value(ctrl_id):
    """Return the deduction value for a control per 32 CFR 170.24, plus its class label."""
    if ctrl_id == SPRS_SSP_REQ:
        return 0, 'SSP requirement (no point value — assessment cannot complete without it)'
    if ctrl_id in SPRS_SPECIAL:
        return 5, 'Partial-credit (5 if not implemented, 3 if partially implemented)'
    if ctrl_id in SPRS_FIVE_BASIC:
        return 5, 'Basic — significant exploitation risk'
    if ctrl_id in SPRS_FIVE_DERIVED:
        return 5, 'Derived — significant exploitation risk'
    if ctrl_id in SPRS_THREE_BASIC:
        return 3, 'Basic — specific and confined effect'
    if ctrl_id in SPRS_THREE_DERIVED:
        return 3, 'Derived — specific and confined effect'
    return 1, 'Derived — limited or indirect effect'

def compute_sprs(ssp):
    """
    Compute estimated SPRS score per 32 CFR 170.24(c)(2).
    MET = Implemented; N/A counts as MET per the rule.
    'Planned to be Implemented' and 'Not Implemented' both count as NOT MET —
    a POA&M is not a substitute for implementation per 170.24(c)(2)(i)(A)(6).
    Partial credit for IA.L2-3.5.3 / SC.L2-3.13.11 comes from ssp['sprs_partial'].
    Returns a dict of score data for the dashboard.
    """
    statuses = ssp.get('controls_status', {})
    partials = ssp.get('sprs_partial', {})   # {ctrl_id: 'partial'|'none'}

    score = 110
    met_count = 0
    na_count = 0
    unmet = []          # list of dicts: id, name, points, status, klass
    ssp_missing = False

    for ctrl in CMMC_CONTROLS:
        cid = ctrl['id']
        # Conservative default: a control never marked counts as NOT MET.
        status = statuses.get(cid, 'Not Implemented')
        is_met = status in ('Implemented', 'Not Applicable')

        if status == 'Not Applicable':
            na_count += 1
        elif is_met:
            met_count += 1

        if is_met:
            continue

        # NOT MET path
        if cid == SPRS_SSP_REQ:
            ssp_missing = True
            unmet.append({'id': cid, 'name': ctrl['name'], 'points': 0,
                          'status': status, 'klass': 'SSP requirement'})
            continue

        if cid in SPRS_SPECIAL:
            deduction = 3 if partials.get(cid) == 'partial' else 5
            label = 'Partially implemented (−3)' if partials.get(cid) == 'partial' else 'Not implemented (−5)'
            score -= deduction
            unmet.append({'id': cid, 'name': ctrl['name'], 'points': deduction,
                          'status': status, 'klass': label})
            continue

        pts, klass = sprs_point_value(cid)
        score -= pts
        unmet.append({'id': cid, 'name': ctrl['name'], 'points': pts,
                      'status': status, 'klass': klass})

    # Sort unmet by highest deduction first
    unmet.sort(key=lambda u: (-u['points'], u['id']))

    marked = len(statuses)
    return {
        'score': score,
        'max_score': 110,
        'min_score': -203,
        'met': met_count,
        'na': na_count,
        'not_met': len(unmet),
        'marked': marked,
        'unmarked': 110 - marked if marked < 110 else 0,
        'unmet': unmet,
        'deducted': 110 - score,
        'ssp_missing': ssp_missing,
        'pct': round(((score + 203) / (110 + 203)) * 100, 1),
    }

# ─────────────────────────────────────────────────────────────
# POA&M — Plan of Action & Milestones
# Structure follows the standard OSC POA&M workbook (17 columns).
# Seeded from poam_seed.DEFAULT_POAM, which maps all 110
# NIST SP 800-171 R2 requirements across 68 POA&M items.
# ─────────────────────────────────────────────────────────────
POAM_STATUSES = ['Not Started', 'In Progress', 'Completed', 'Risk Accepted', 'Deferred']
POAM_RISK_LEVELS = ['Low', 'Med', 'High']
POAM_IDENTIFIERS = ['Gap Analysis', 'Security Configuration Review', 'Vulnerability Scan',
                    'Penetration Test', 'Internal Audit', 'Self-Assessment', 'C3PAO Assessment',
                    'Incident', 'Other']
POAM_COST_STATES = ['Funded', 'Unfunded', 'Partially Funded']
POAM_MILESTONES = ['Confirmed', 'Confirmed Configuration', 'Confirmed Documentation',
                   'Documentation Delivered', 'In Remediation', 'Pending Validation']

POAM_FIELDS = ['poam_id', 'cmmc_ref', 'nist_ref', 'owner', 'deficiency_details',
               'deficiency_date', 'deficiency_identifier', 'remediation', 'estimated_cost',
               'date_complete', 'milestone', 'ttp_identifier', 'risk_likelihood',
               'risk_impact', 'risk_level', 'status', 'milestones_notes']

def load_poam():
    if POAM_FILE.exists():
        with open(POAM_FILE) as f:
            data = json.load(f)
        data.setdefault('items', [])
        for it in data['items']:
            it.setdefault('id', str(uuid.uuid4()))
            for fld in POAM_FIELDS:
                it.setdefault(fld, '')
        return data
    # First run — seed from the template
    try:
        from poam_seed import DEFAULT_POAM
        import copy
        items = copy.deepcopy(DEFAULT_POAM)
    except Exception:
        items = []
    for it in items:
        it['id'] = str(uuid.uuid4())
        for fld in POAM_FIELDS:
            it.setdefault(fld, '')
    data = {'items': items}
    save_poam(data)
    return data

def save_poam(data):
    with open(POAM_FILE, 'w') as f:
        json.dump(data, f, indent=2)

def _nist_refs_in(text):
    """Pull NIST SP 800-171 requirement numbers (3.x.y) out of a reference cell."""
    import re as _re
    return _re.findall(r'3\.\d+\.\d+', str(text or ''))

def poam_coverage(poam_items):
    """
    Audit which of the 110 NIST SP 800-171 R2 requirements are represented
    in the POA&M, and which POA&M item(s) each maps to.
    """
    ref_to_ctrl = {c['id'].split('-', 1)[1]: c for c in CMMC_CONTROLS}
    covered = {}
    for it in poam_items:
        for ref in _nist_refs_in(it.get('nist_ref')):
            covered.setdefault(ref, []).append(it.get('poam_id', ''))
    missing = []
    for ref, ctrl in ref_to_ctrl.items():
        if ref not in covered:
            missing.append({'nist_ref': ref, 'ctrl_id': ctrl['id'], 'name': ctrl['name']})
    missing.sort(key=lambda m: [int(x) for x in m['nist_ref'].split('.')])
    return {
        'total_reqs': len(ref_to_ctrl),
        'covered': len([r for r in covered if r in ref_to_ctrl]),
        'missing': missing,
        'complete': len(missing) == 0,
        'map': covered,
    }

def poam_summary(poam_items, ssp):
    """
    Roll up POA&M status and tie it to the SPRS score: how many points are
    currently being lost by requirements that have an open POA&M item.
    """
    statuses = ssp.get('controls_status', {})
    ref_to_ctrl = {c['id'].split('-', 1)[1]: c for c in CMMC_CONTROLS}

    by_status = {s: 0 for s in POAM_STATUSES}
    by_risk = {r: 0 for r in POAM_RISK_LEVELS}
    open_items = 0
    recoverable = 0
    seen_refs = set()

    for it in poam_items:
        st = it.get('status') or 'Not Started'
        by_status[st] = by_status.get(st, 0) + 1
        rl = it.get('risk_level')
        if rl in by_risk:
            by_risk[rl] += 1
        is_open = st not in ('Completed', 'Risk Accepted')
        if is_open:
            open_items += 1
        # SPRS points recoverable: sum deductions for Not Met controls on open items
        for ref in _nist_refs_in(it.get('nist_ref')):
            ctrl = ref_to_ctrl.get(ref)
            if not ctrl or ref in seen_refs:
                continue
            cstatus = statuses.get(ctrl['id'], 'Not Implemented')
            if cstatus not in ('Implemented', 'Not Applicable') and is_open:
                seen_refs.add(ref)
                pts, _ = sprs_point_value(ctrl['id'])
                if ctrl['id'] in SPRS_SPECIAL:
                    pts = 3 if ssp.get('sprs_partial', {}).get(ctrl['id']) == 'partial' else 5
                recoverable += pts

    return {
        'total': len(poam_items),
        'open': open_items,
        'closed': len(poam_items) - open_items,
        'by_status': by_status,
        'by_risk': by_risk,
        'sprs_recoverable': recoverable,
    }

def poam_control_status(item, ssp):
    """For a POA&M item, summarise the live control status of its NIST refs."""
    statuses = ssp.get('controls_status', {})
    ref_to_ctrl = {c['id'].split('-', 1)[1]: c for c in CMMC_CONTROLS}
    met = notmet = 0
    for ref in _nist_refs_in(item.get('nist_ref')):
        ctrl = ref_to_ctrl.get(ref)
        if not ctrl:
            continue
        st = statuses.get(ctrl['id'], 'Not Implemented')
        if st in ('Implemented', 'Not Applicable'):
            met += 1
        else:
            notmet += 1
    return {'met': met, 'not_met': notmet}

# ─────────────────────────────────────────────────────────────
# Gap Analysis — pre-assessment discovery checklist
# One row per NIST SP 800-171 R2 assessment objective (320 total across
# the 110 practices), sourced from the client's CMMC L2 Gap Analysis Full
# Environment workbook (see gap_seed.py for provenance). This runs ahead
# of the Controls Status page: work the questions with the client/
# stakeholder, capture Status/Evidence/Gap here, then set the resulting
# implementation status on the Controls page and (for anything Not Met)
# a POA&M item.
# ─────────────────────────────────────────────────────────────
GAP_STATUSES = ['Met', 'Partially Met', 'Not Met', 'N/A']

def gap_summary(ssp):
    """Roll up Gap Analysis status counts across all 320 assessment objectives."""
    data = ssp.get('gap_analysis', {})
    counts = {s: 0 for s in GAP_STATUSES}
    unanswered = 0
    total = 0
    for ctrl in CMMC_CONTROLS:
        for letter, _ in ctrl['objectives']:
            total += 1
            entry = (data.get(ctrl['id']) or {}).get(letter) or {}
            st = entry.get('status', '')
            if st in counts:
                counts[st] += 1
            else:
                unanswered += 1
    answered = total - unanswered
    return {
        'total': total,
        'answered': answered,
        'unanswered': unanswered,
        'counts': counts,
        'pct_answered': round(answered / total * 100, 1) if total else 0,
        'gaps': counts['Partially Met'] + counts['Not Met'],
    }

def sync_gap_deficiencies_to_poam(ssp, poam_data):
    """
    Push Gap Analysis 'Gap/Deficiency' text into the matching POA&M
    item(s)' Deficiency Details field, keyed by the shared NIST SP
    800-171 requirement number (e.g. AC.L2-3.1.1 -> '3.1.1'). Only
    touches controls that currently have at least one non-blank
    Gap/Deficiency entry — a control with nothing recorded in Gap
    Analysis is left alone, so this never silently blanks out
    something typed directly into POA&M. Returns the number of POA&M
    items updated.
    """
    gap = ssp.get('gap_analysis', {})
    items = poam_data.get('items', [])
    updated = 0
    for ctrl_id, objectives in gap.items():
        if '-' not in ctrl_id:
            continue
        nist_ref = ctrl_id.split('-', 1)[1]
        parts = [(letter, (objectives.get(letter) or {}).get('gap_deficiency', '').strip())
                 for letter in sorted(objectives.keys())]
        parts = [(l, t) for l, t in parts if t]
        if not parts:
            continue
        combined = parts[0][1] if len(parts) == 1 else '\n'.join(f"[{l}] {t}" for l, t in parts)
        for item in items:
            if nist_ref in _nist_refs_in(item.get('nist_ref')) and item.get('deficiency_details', '') != combined:
                item['deficiency_details'] = combined
                updated += 1
    return updated

# ─────────────────────────────────────────────────────────────
# Security & Risk Analysis — NIST SP 800-53-based assessment
# A distinct framework from the CMMC L2 / 800-171 tooling above: 222
# items across 3 control-set worksheets (see sra_seed.py for
# provenance), rolled up by NIST 800-53 control family for the
# percent-met-by-domain report the export produces.
# ─────────────────────────────────────────────────────────────
SRA_STATUSES = ['Not Assessed', 'Meets', 'Does Not Meet', 'Not Applicable']
SRA_FIELDS = ['observations', 'interview', 'test', 'criteria', 'notes']
SRA_ITEM_IDS = {it['id'] for it in SRA_ITEMS}

def load_sra():
    if SRA_FILE.exists():
        with open(SRA_FILE) as f:
            data = json.load(f)
        data.setdefault('items', {})
        return data
    return {'items': {}}

def save_sra(data):
    with open(SRA_FILE, 'w') as f:
        json.dump(data, f, indent=2)

def sra_grouped():
    """
    Nest SRA_ITEMS as family -> [{ref, ref_name, items:[...]}], in the
    families' canonical order, preserving each item's original row order
    within its control (already contiguous in SRA_ITEMS).
    """
    fam_map = {}
    for it in SRA_ITEMS:
        fam = fam_map.setdefault(it['family'], {
            'code': it['family'], 'name': SRA_FAMILIES.get(it['family'], it['family']), 'controls': []
        })
        controls = fam['controls']
        if not controls or controls[-1]['ref'] != it['ref']:
            # 'entries' (not 'items') — a dict's own .items() method would
            # shadow a Jinja `ctrl.items` attribute lookup in the template.
            controls.append({'ref': it['ref'], 'ref_name': it['ref_name'], 'entries': []})
        controls[-1]['entries'].append(it)
    return [fam_map[code] for code in SRA_FAMILIES if code in fam_map]

def sra_rollup(items_subset, sra_data):
    """Status counts + percent-met for a subset of SRA_ITEMS. Not Applicable
    items are excluded from the percent-met denominator; Not Assessed items
    count against it (conservative — matches this app's SPRS/Gap Analysis
    convention of never crediting an unreviewed item)."""
    counts = {s: 0 for s in SRA_STATUSES}
    for it in items_subset:
        entry = sra_data.get('items', {}).get(it['id'], {})
        st = entry.get('criteria') or 'Not Assessed'
        counts[st] = counts.get(st, 0) + 1
    total = len(items_subset)
    denom = total - counts['Not Applicable']
    pct_met = round(counts['Meets'] / denom * 100, 1) if denom > 0 else None
    return {'total': total, 'counts': counts, 'pct_met': pct_met}

def sra_summary(sra_data):
    overall = sra_rollup(SRA_ITEMS, sra_data)
    by_family = []
    for code, name in SRA_FAMILIES.items():
        fam_items = [it for it in SRA_ITEMS if it['family'] == code]
        if not fam_items:
            continue
        by_family.append({'code': code, 'name': name, **sra_rollup(fam_items, sra_data)})
    by_sheet = []
    for sheet in SRA_SHEETS:
        sheet_items = [it for it in SRA_ITEMS if it['sheet'] == sheet]
        by_sheet.append({'sheet': sheet, 'label': SRA_SHEET_LABELS[sheet], **sra_rollup(sheet_items, sra_data)})
    return {'overall': overall, 'by_family': by_family, 'by_sheet': by_sheet}

def sra_deficiencies(sra_data):
    """Items marked Does Not Meet, for the Key Remediation Points report section."""
    out = []
    for it in SRA_ITEMS:
        entry = sra_data.get('items', {}).get(it['id'], {})
        if (entry.get('criteria') or 'Not Assessed') == 'Does Not Meet':
            out.append({**it, **{f: entry.get(f, '') for f in SRA_FIELDS}})
    out.sort(key=lambda x: (x['family'], x['ref'] or '', x['id']))
    return out

ASSET_REFERENCE_MAP = {    # AC.L2-3.1.1 — Authorized Access Control
    "AC.L2-3.1.1|a": ("users", "Authorized Users"),
    "AC.L2-3.1.1|b": ("users", "Authorized Users"),
    "AC.L2-3.1.1|c": ("it-assets", "IT Assets"),
    "AC.L2-3.1.1|d": ("users", "Authorized Users"),
    "AC.L2-3.1.1|e": ("users", "Authorized Users"),
    "AC.L2-3.1.1|f": ("it-assets", "IT Assets"),
    # AC.L2-3.1.2 — Transaction & Function Control
    "AC.L2-3.1.2|a": ("users", "Authorized Users"),
    "AC.L2-3.1.2|b": ("users", "Authorized Users"),
    # AC.L2-3.1.3 — Control CUI Flow
    "AC.L2-3.1.3|c": ("it-assets", "IT Assets"),
    # AC.L2-3.1.5 — Least Privilege
    "AC.L2-3.1.5|a": ("users", "Authorized Users (Privileged)"),
    "AC.L2-3.1.5|b": ("users", "Authorized Users (Privileged)"),
    # AC.L2-3.1.6 — Non-Privileged Account Use
    "AC.L2-3.1.6|b": ("users", "Authorized Users (Privileged)"),
    # AC.L2-3.1.7 — Privileged Functions
    "AC.L2-3.1.7|b": ("users", "Authorized Users"),
    # AC.L2-3.1.12 — Control Remote Access
    "AC.L2-3.1.12|b": ("ports", "Authorized Ports & Protocols"),
    # AC.L2-3.1.14 — Remote Access Routing
    "AC.L2-3.1.14|a": ("ports", "Authorized Ports & Protocols"),
    # AC.L2-3.1.15 — Privileged Remote Access
    "AC.L2-3.1.15|a": ("users", "Authorized Users (Privileged)"),
    "AC.L2-3.1.15|b": ("users", "Authorized Users (Privileged)"),
    # AC.L2-3.1.16 — Wireless Access Authorization
    "AC.L2-3.1.16|a": ("it-assets", "IT Assets"),
    # AC.L2-3.1.18 — Mobile Device Connection
    "AC.L2-3.1.18|a": ("it-assets", "IT Assets"),
    # AC.L2-3.1.19 — Encrypt CUI on Mobile
    "AC.L2-3.1.19|a": ("it-assets", "IT Assets"),
    # AC.L2-3.1.20 — External Connections
    "AC.L2-3.1.20|a": ("it-assets", "IT Assets"),
    "AC.L2-3.1.20|b": ("serverapps", "Authorized Server Apps"),
    # AC.L2-3.1.21 — Portable Storage Use
    "AC.L2-3.1.21|a": ("it-assets", "IT Assets"),
    # AC.L2-3.1.22 — Control Public Information
    "AC.L2-3.1.22|a": ("users", "Authorized Users"),
    # AT.L2-3.2.2 — Role-Based Training
    "AT.L2-3.2.2|b": ("users", "Authorized Users"),
    # AU.L2-3.3.4 — Audit Failure Alerting
    "AU.L2-3.3.4|a": ("users", "Authorized Users"),
    # AU.L2-3.3.9 — Audit Management
    "AU.L2-3.3.9|a": ("users", "Authorized Users (Privileged)"),
    # CM.L2-3.4.1 — System Baselining
    "CM.L2-3.4.1|d": ("it-assets", "IT Assets"),
    "CM.L2-3.4.1|e": ("it-assets", "IT Assets"),
    "CM.L2-3.4.1|f": ("it-assets", "IT Assets"),
    # CM.L2-3.4.5 — Access Restrictions for Change
    "CM.L2-3.4.5|a": ("phys-assets", "Physical Security Assets"),
    "CM.L2-3.4.5|b": ("phys-assets", "Physical Security Assets"),
    # CM.L2-3.4.7 — Nonessential Functionality
    "CM.L2-3.4.7|a": ("serverapps", "Authorized Server Apps"),
    "CM.L2-3.4.7|b": ("blacklist", "Workstation App Blacklist"),
    "CM.L2-3.4.7|c": ("blacklist", "Workstation App Blacklist"),
    "CM.L2-3.4.7|d": ("processes", "Authorized System Processes"),
    "CM.L2-3.4.7|e": ("processes", "Authorized System Processes"),
    "CM.L2-3.4.7|g": ("ports", "Authorized Ports & Protocols"),
    "CM.L2-3.4.7|h": ("ports", "Authorized Ports & Protocols"),
    "CM.L2-3.4.7|i": ("ports", "Authorized Ports & Protocols"),
    "CM.L2-3.4.7|j": ("ports", "Authorized Ports & Protocols"),
    "CM.L2-3.4.7|k": ("ports", "Authorized Ports & Protocols"),
    "CM.L2-3.4.7|m": ("serverapps", "Authorized Server Apps"),
    "CM.L2-3.4.7|n": ("serverapps", "Authorized Server Apps"),
    # CM.L2-3.4.8 — Application Execution Policy
    "CM.L2-3.4.8|b": ("serverapps", "Authorized Server Apps / Workstation App Blacklist"),
    # CM.L2-3.4.9 — User-Installed Software
    "CM.L2-3.4.9|a": ("blacklist", "Workstation App Blacklist"),
    # IA.L2-3.5.1 — Identification
    "IA.L2-3.5.1|a": ("users", "Authorized Users"),
    "IA.L2-3.5.1|b": ("users", "Authorized Users"),
    "IA.L2-3.5.1|c": ("it-assets", "IT Assets"),
    # IA.L2-3.5.3 — Multifactor Authentication
    "IA.L2-3.5.3|a": ("users", "Authorized Users (Privileged)"),
    # IA.L2-3.5.4 — Replay-Resistant Authentication
    "IA.L2-3.5.4|a": ("users", "Authorized Users (Privileged)"),
    # MP.L2-3.8.2 — Media Access
    "MP.L2-3.8.2|a": ("users", "Authorized Users"),
    # PS.L2-3.9.1 — Screen Individuals
    "PS.L2-3.9.1|a": ("users", "Authorized Users"),
    # PE.L2-3.10.1 — Limit Physical Access
    "PE.L2-3.10.1|a": ("users", "Authorized Users"),
    # PE.L2-3.10.5 — Manage Physical Access Devices
    "PE.L2-3.10.5|a": ("phys-assets", "Physical Security Assets"),
    # SC.L2-3.13.5 — Public-Access System Separation
    "SC.L2-3.13.5|a": ("it-assets", "IT Assets"),
    # SC.L2-3.13.6 — Network Communication by Exception
    "SC.L2-3.13.6|a": ("ports", "Authorized Ports & Protocols"),
    "SC.L2-3.13.6|b": ("ports", "Authorized Ports & Protocols"),
    # SC.L2-3.13.12 — Collaborative Computing
    "SC.L2-3.13.12|a": ("it-assets", "IT Assets"),
    # SC.L2-3.13.13 — Mobile Code
    "SC.L2-3.13.13|a": ("serverapps", "Authorized Server Apps"),
    # SC.L2-3.13.14 — VoIP
    "SC.L2-3.13.14|a": ("it-assets", "IT Assets"),
    # SI.L2-3.14.7 — Identify Unauthorized Use
    "SI.L2-3.14.7|a": ("users", "Authorized Users"),
}

# ─────────────────────────────────────────────────────────────
# Data Persistence
# ─────────────────────────────────────────────────────────────
def load_ssp():
    if SSP_FILE.exists():
        with open(SSP_FILE) as f:
            data = json.load(f)
        # Migrate: ensure newer keys exist on older saved files
        defaults = get_default_ssp()
        for key, val in defaults.items():
            if key not in data:
                data[key] = val
        if 'images' in data:
            for img_key, img_val in defaults['images'].items():
                if img_key not in data['images']:
                    data['images'][img_key] = img_val
            # Migrate: network/physical/cui_flow diagrams used to be a single
            # filename (or None); they're now a list so a site can have more
            # than one. Wrap any old-format value so nothing already uploaded
            # gets dropped.
            for img_key in ('network_diagram', 'physical_diagram', 'cui_flow_diagram'):
                val = data['images'].get(img_key)
                if isinstance(val, str):
                    data['images'][img_key] = [val]
                elif val is None:
                    data['images'][img_key] = []
        return data
    return get_default_ssp()

def save_ssp(data):
    with open(SSP_FILE, 'w') as f:
        json.dump(data, f, indent=2)

DEFAULT_BLACKLIST = [
    {"id": "bl-p2p",      "name": "Peer-to-Peer Applications",           "justification": "Increases risk of bypassing security controls, distribution of malware, and data exfiltration.",       "apps": ["BitTorrent","uTorrent","qBittorrent"]},
    {"id": "bl-anon",     "name": "Anonymous Network Applications",       "justification": "Bypasses monitoring capabilities.",                                                                       "apps": ["Tor Browser","I2P","Psiphon"]},
    {"id": "bl-remote",   "name": "Unapproved Remote Access Applications","justification": "Allows for backdoors for threat actors and increases likelihood of lateral movement.",                  "apps": ["TeamViewer","AnyDesk","LogMeIn"]},
    {"id": "bl-crypto",   "name": "Crypto Mining Software",              "justification": "Impacts the available company resources and workstation performance.",                                   "apps": ["XMRig","cpuminer","claymore"]},
    {"id": "bl-creds",    "name": "Credential Theft Tools",              "justification": "May be used to steal credentials and escalate privileges.",                                              "apps": ["Mimikatz","Hashcat","LaZagne"]},
    {"id": "bl-vpn",      "name": "Non-Approved VPNs",                   "justification": "Increases risk of unmonitored data exfiltration.",                                                       "apps": ["ExpressVPN","NordVPN","ProtonVPN"]},
    {"id": "bl-nonbiz",   "name": "Non-Business Applications (adjustable as needed)", "justification": "", "apps": ["Spotify","Netflix","TikTok"]},
]

def load_assets():
    if ASSET_FILE.exists():
        with open(ASSET_FILE) as f:
            data = json.load(f)
        # Migrate: ensure newer sections exist on older saved files
        for key in ['assets', 'physical_security', 'authorized_users',
                    'server_apps', 'mobile_whitelist', 'system_processes', 'ports_protocols']:
            data.setdefault(key, [])
        # Migrate blacklist: seed from defaults if missing
        if 'blacklist_categories' not in data:
            import copy
            data['blacklist_categories'] = copy.deepcopy(DEFAULT_BLACKLIST)
        if 'blacklist_tech_impl' not in data:
            data['blacklist_tech_impl'] = "The listed applications are restricted via Defender for Endpoint within Settings > Endpoint > Indicators, and exported AppLocker controls deployed via Intune."
        if 'blacklist_notice' not in data:
            data['blacklist_notice'] = "While a blacklist of known high-risk or unwanted applications is maintained, applications should not be downloaded without approval from IT Security, IT Support, and the Security Program Coordinator. The absence of an application from the blacklist does not mean it is permitted without change approval."
        # Migrate: ensure newer per-user privileged-access fields exist
        user_defaults = {
            'is_privileged': False, 'privileged_server': [], 'privileged_firewall': [],
            'privileged_network_device': [], 'privileged_ms365': [], 'pim_privileges': [],
        }
        for u in data.get('authorized_users', []):
            for k, v in user_defaults.items():
                u.setdefault(k, v)
        # Migrate: ensure newer server app / mobile whitelist fields exist
        app_defaults = {'app_type': '', 'deployment_model': '', 'fedramp_approved': ''}
        for a in data.get('server_apps', []):
            for k, v in app_defaults.items():
                a.setdefault(k, v)
        for a in data.get('mobile_whitelist', []):
            for k, v in app_defaults.items():
                a.setdefault(k, v)
        return data
    import copy
    return {
        "assets": [],
        "physical_security": [],
        "authorized_users": [],
        "server_apps": [],
        "mobile_whitelist": [],
        "system_processes": [],
        "ports_protocols": [],
        "blacklist_categories": copy.deepcopy(DEFAULT_BLACKLIST),
        "blacklist_tech_impl": "The listed applications are restricted via Defender for Endpoint within Settings > Endpoint > Indicators, and exported AppLocker controls deployed via Intune.",
        "blacklist_notice": "While a blacklist of known high-risk or unwanted applications is maintained, applications should not be downloaded without approval from IT Security, IT Support, and the Security Program Coordinator. The absence of an application from the blacklist does not mean it is permitted without change approval.",
    }

def save_assets(data):
    with open(ASSET_FILE, 'w') as f:
        json.dump(data, f, indent=2)

def get_default_ssp():
    return {
        "system_name": "",
        "system_type": "Multi-User Standalone (MUSA, M365) System, Wide Area Network (WAN)",
        "system_categorization": "CUI, FCI",
        "system_uid": "IT Infrastructure",
        "version": "1.0",
        "last_updated": datetime.now().strftime("%m.%d.%Y"),
        "org_name": "",
        "org_abbr": "",
        "mssp_name": "",
        "prepared_by": "",
        "system_owner": {"name": "", "title": "", "email": "", "phone": ""},
        "info_owner": {"name": "", "title": "", "email": "", "phone": ""},
        "sso": {"name": "", "title": "", "email": "", "phone": ""},
        "end_users": "",
        "privileged_users": "",
        "cui_categories": "",
        "nara_cui_intro": "As defined by the U.S. National Archives and Records Administration, [OSC] processes, "
                          "transmits or stores the following categories of CUI",
        "nara_cui_categories": [],
        "locations": [],
        "main_office_site": "",
        "home_office_site": "",
        "system_environment": "The following is a detailed topology narrative and graphic that clearly depicts the "
                              "system boundaries, system interconnections, key devices, and physical / logical "
                              "security parameters.",
        "cui_data_flow": "The following is a detailed data flow narrative and graphic that clearly depicts the flow "
                         "of CUI electronic media that is processed, stored, and transmitted into the CUI system.",
        "physical_security_desc": "",   # legacy single-narrative field — superseded by physical_security_by_location
        "physical_security_by_location": {},   # {location_name: narrative}
        "network_desc": "",             # legacy single-narrative field — superseded by spa_network_by_location
        "spa_network_by_location": {},         # {location_name: narrative} — firewall/IPSEC per site
        "server_desc": "The [OSC] does/does not utilize servers within the operational system environment.",
        "server_types": [],             # list of strings, e.g. "Domain Controller (Hyper-V)"
        "workstation_desc": "The [OSC] standardized its operating system to Windows 11 Professional. The [OSC] "
                            "employees are issued laptops [or other IT systems]. All systems are centrally managed "
                            "and configured to support secure processing, storage, and transmission of CUI material. "
                            "Only authorized, company-managed systems that are Microsoft Intune joined are permitted "
                            "to process, store, or transmit CUI material. Personal devices are strictly prohibited "
                            "from accessing CUI.\n\n"
                            "All systems are configured using a standardized secure baseline configuration enforced "
                            "through Microsoft Intune. Full-disk encryption (BitLocker) is enabled, USB storage is "
                            "restricted, host-based firewalls are active, and automatic inactive session locking is "
                            "enforced.\n\n"
                            "Users authenticate using unique Microsoft Entra ID credentials with MFA enforced for "
                            "privileged and remote access. Local administrator access is restricted to authorized IT "
                            "personnel.\n\n"
                            "All systems utilize endpoint detection and response (EDR) software with real-time "
                            "monitoring and centralized alerting. Security events are logged and reviewed in "
                            "accordance with the incident response program. Operating system and application "
                            "updates are automatically deployed, with critical patches expedited. Compliance is "
                            "centrally monitored by IT administrators.\n\n"
                            "Systems are used within controlled office environments or approved remote locations. "
                            "Users are required to secure devices when unattended and report lost or stolen devices "
                            "in accordance with the documented organizational policy.",
        "crma_mobile_desc": "The [OSC] employs mobile devices within their environment, managed via Intune "
                            "Configuration policy. The in-scope mobile device program includes company-owned and "
                            "authorized BYOD smartphones used by personnel to access email and applications "
                            "containing CUI material.\n\n"
                            "Only authorized mobile devices enrolled in the organization's Mobile Device Management "
                            "(MDM) solution are permitted to access CUI material. The MDM solution enforces security "
                            "configuration policies, monitors compliance, and restricts access to sensitive business "
                            "resources based on device health. Unauthorized and/or unenrolled devices are "
                            "blocked.\n\n"
                            "Access to CUI material from mobile devices requires multi-factor authentication and is "
                            "permitted only from compliant devices. Intune security configuration policies block "
                            "access from non-compliant or jailbroken/rooted devices.\n\n"
                            "Mobile devices enforce native operating system encryption for data at rest. Any CUI "
                            "material transmitted to or from mobile devices is protected using FIPS-validated "
                            "cryptographic modules. CUI material accessed on mobile devices is restricted to managed "
                            "applications and encrypted containers. Downloading CUI material to unmanaged "
                            "applications or personal storage locations is prohibited and technically blocked.\n\n"
                            "Lost or stolen mobile devices must be reported immediately in accordance with the "
                            "documented organizational policy. IT administrators can remotely lock or wipe devices "
                            "to prevent unauthorized access to CUI material. Mobile device security events and "
                            "compliance status are logged and monitored centrally. Non-compliant devices are "
                            "automatically restricted from accessing CUI material.\n\n"
                            "Users are responsible for maintaining physical control of mobile devices and "
                            "preventing unauthorized viewing or access to CUI material while in use.\n\n"
                            "Mobile device use is governed by the System Security Plan, Incident Response Program, "
                            "Mobile Device Management Policy, and the Media Protection Policy.",
        "specialized_assets_desc": "",   # free text — Specialized Assets (SA), if any (GFE, IoT, OT, etc.)
        "spa_wireless_desc": "The [OSC] does/does not authorize wireless access points at the Home Office, "
                             "Satellite Office, and/or Remote Office.\n\n"
                             "Access to the corporate network is only permitted by [wireless access points and/or a "
                             "physical ethernet connection].\n\n"
                             "Authentication to the network is managed by Microsoft Entra ID and Microsoft Intune.",
        "spa_usb_desc": "The [OSC] authorizes and whitelists the use of USB portable storage devices on the CUI "
                        "system.",
        "cloud_apps_desc": "The [OSC] utilizes Microsoft 365 GCC[-High] G5 to process, transmit and store CUI "
                           "material. Additionally, Microsoft 365 is used to provide various security capabilities "
                           "including:\n\n"
                           "- Multifactor Authentication\n"
                           "- FIPS 140 Validated Encryption\n"
                           "- Information Classification\n"
                           "- Security Configuration Management",
        "out_of_scope_assets_desc": "", # free text — Out-of-Scope Assets (OOS), if any

        # ── Controls (Administrative Controls / Technical Controls narrative section) ──
        "controls_intro": "The [OSC] utilizes the following controls to protect CUI material.",
        "admin_controls_intro": "The following controls are procedural and designed to promote workforce security.",
        "policies_intro": "The [OSC] implements a comprehensive policy set of documentation for Information "
                          "Technology and user (employee) governance:",
        "policies_list": "ITP-01 – Information Protection Policy\n"
                         "ITP-02 – Network Protection Policy\n"
                         "ITP-03 – System (Server) Security Policy\n"
                         "ITP-04 – Endpoint Security Policy\n"
                         "ITP-05 – Application Security Policy\n"
                         "ITP-06 – Configuration Management Policy\n"
                         "ITP-07 – Security Governance Policy\n"
                         "ITP-08 – Incident Management Policy\n"
                         "ITP-09 – Data Center Policy\n"
                         "UP-01 – Information Security Policy for End Users\n"
                         "UP-02 – Records Retention Policy",
        "standards_list": "ITS-01-01 - Workplace Security Standard\n"
                          "ITS-01-02 – Encryption Standard\n"
                          "ITS-01-03 – Access Control Standard\n"
                          "ITS-01-04 – Account Management Standard\n"
                          "ITS-01-06 – Media Handling Standard\n"
                          "ITS-01-07 – Password Vault Security\n"
                          "ITS-01-09 – Bitwarden Security\n"
                          "ITS-02-03 - Wireless Security Standard\n"
                          "ITS-02-04 - FortiGate Security Standard\n"
                          "ITS-03-01 - Windows Server 2022 Security\n"
                          "ITS-03-02 - Ubuntu 20.04 Security\n"
                          "ITS-03-03 - Red Hat (RHEL) 8 Security\n"
                          "ITS-03-10 - Ubuntu 18.04 Security\n"
                          "ITS-03-12 – Windows Server 2019 Security\n"
                          "ITS-04-02 - Windows 11 Security Standard\n"
                          "ITS-04-03 – Apple iOS Security\n"
                          "ITS-04-04 – Google Android Security\n"
                          "ITS-04-05 - Microsoft 365 Defender Security\n"
                          "ITS-04-09 – Mobile Device Management Standard\n"
                          "ITS-04-10 – Windows 10 Security\n"
                          "ITS-04-11 - Microsoft 365 Intune Standard\n"
                          "ITS-04-12 - Microsoft 365 SharePoint & OneDrive Standard\n"
                          "ITS-06-01 – System Update Security Standard\n"
                          "ITS-06-02 – Vulnerability Management Standard\n"
                          "ITS-06-03 – Change Management Standard\n"
                          "ITS-06-04 - Microsoft 365 Entra ID Security Standard\n"
                          "ITS-06-05 - Microsoft 365 Teams Security Standard\n"
                          "ITS-07-01 – Asset Management Standard\n"
                          "ITS-07-02 – System Logging Standard\n"
                          "ITS-07-03 – Architecture Security Standard\n"
                          "ITS-07-04 - Compliance Management Standard\n"
                          "ITS-09-01 - Data Center Physical Security Standard\n"
                          "ITS-09-02 - Data Center Environment Standard\n"
                          "ITS-09-03 - Data Center Power Continuity Standard\n"
                          "ITS-09-04 - Data Center Monitoring Standard\n"
                          "ITS-09-05 - Data Center Fire Suppression Standard\n"
                          "ITS-09-06 - Data Center Training Standard\n"
                          "US-01-01 – Sensitive Information Protection Standard\n"
                          "[OSC] Incident Response Plan",
        "procedures_list": "ITPC-01-01 - System Access Request Procedure\n"
                           "ITPC-01-02 - Password Management Procedure\n"
                           "ITPC-01-03 - System Access Modification\n"
                           "ITPC-01-04 - System Access Termination\n"
                           "ITPC-01-05 - Media Sanitization Procedure\n"
                           "ITPC-01-10 - CUI Media Handling Procedure\n"
                           "ITPC-01-13 - Visitor Access Procedure\n"
                           "ITPC-02-01 – Third-party Connectivity\n"
                           "ITPC-04-01 - Windows Intune Enrollment\n"
                           "ITPC-04-02 - Apple iOS Enrollment\n"
                           "ITPC-04-03 - Android Intune Enrollment\n"
                           "UPC-01-01 - Internet Posting Approval\n"
                           "UPC-01-02 - VPN Access",
        "access_control_desc": "The [OSC] controls access through Microsoft 365 security groups based on "
                               "pre-defined roles and access control lists. All users have unique identifiers. "
                               "CUI material is segregated into SharePoint folders where access is granted on a "
                               "need-to-know basis.\n\n"
                               "Systems that do not use Microsoft Azure SSO for an authentication source are "
                               "configured to provide authorized users the minimum level of access required to "
                               "perform job functions.\n\n"
                               "The [OSC] enforces least-privilege access across local systems and backup "
                               "solutions. Privileged access to these systems is limited to authorized Sysadmins, "
                               "with access defined by role, responsibility, business functionality, and "
                               "periodically reviewed for validity.",
        "asset_mgmt_desc": "InfoDefense compiles, reviews, and validates the Asset Management Program as a part "
                           "of [OSC]’s CyberSecure 360 subscription. Computer hardware inventories and "
                           "software assets are reported, tracked, and managed in the CyberSecure 360 GRC "
                           "Management platform. These reports support threat and vulnerability analysis, and "
                           "Risk & Security Analysis and Management.",
        "audit_assessment_desc": "InfoDefense performs Security and Risk analysis of the [OSC] system, to "
                                 "include administrative, technical, and physical controls protecting the "
                                 "operational system environment that processes, stores, and transmits CUI "
                                 "material. Along with periodic review, Security and Risk analysis is performed "
                                 "after a major system change and/or upon a security incident affecting the "
                                 "system.\n\n"
                                 "Assessed controls include, but are not limited to authentication, access "
                                 "control, encryption, vulnerability management, security monitoring, network "
                                 "security, malware protection and the organization’s documented policies, "
                                 "standards, procedures, and playbooks.",
        "authentication_desc": "The [OSC] leverages Microsoft 365 Entra ID GCC(-High) as its primary "
                               "authentication provider. Microsoft Entra ID SSO is also used for remote access "
                               "VPN authentication.\n\n"
                               "Windows workstations are Microsoft Intune joined to facilitate management and "
                               "authentication. Each workstation utilizes Microsoft Hello for Business to "
                               "provide local authentication, leveraging a separate PIN/Password coupled with "
                               "keys retrieved from the workstation’s Trusted Platform Module (TPM) to "
                               "authenticate to Microsoft Azure.\n\n"
                               "One Time Password or FIDO MFA is utilized for mobile device and web access to "
                               "Microsoft 365 as well as administrative functionality, IPSEC VPN, maintenance, "
                               "and other types of privileged or remote access.",
        "compliance_mgmt_desc": "InfoDefense performs Compliance Management as a part of [OSC]’s "
                                "CyberSecure 360 subscription. Reminders for the execution of compliance-related "
                                "processes are emailed to organizational system owners at pre-defined intervals. "
                                "When the compliance activity is performed, artifacts (i.e., body of evidence) "
                                "are uploaded to the CyberSecure 360 Compliance Management platform. Compliance "
                                "reports are generated periodically and delivered to [OSC] stakeholders to "
                                "ensure continuous compliance.",
        "incident_response_desc": "InfoDefense manages the Cyber Incident Response Program as a part of "
                                  "[OSC]’s CyberSecure 360 subscription. The [OSC] implements the Incident "
                                  "Response Program to includes a comprehensive incident response plan, "
                                  "procedures, playbooks, periodic training exercises, and rapidly deployable "
                                  "incident response tools to mitigate issues. The incident response plan "
                                  "includes user activities, stakeholder activities, and law enforcement contact "
                                  "information.",
        "personnel_security_desc": "The [OSC] manages and performs background checks for all employees as a "
                                   "prerequisite to accessing CUI material. These background checks include "
                                   "federal, state, and local criminal history, terror watch list, and other "
                                   "criminal background screening as defined by the [the chosen criminal "
                                   "background screening portal, application, or service]. The [OSC] utilizes "
                                   "procedural System Access Request and System Access Termination guidance to "
                                   "grant access to system resources to onboard and off-board [OSC] employees. "
                                   "The [OSC] employee handbook accounts for all other organizational defined "
                                   "employment-related personnel security requirements.",
        "risk_analysis_admin_desc": "InfoDefense performs Security & Risk analysis as a part of [OSC]’s "
                                    "CyberSecure 360 subscription. Security & Risk are continuously evaluated "
                                    "upon a significant change to the operational system environment and/or "
                                    "upon discovery of a cybersecurity incident. Periodically, the [OSC] risk "
                                    "posture is reviewed and adjusted in response to emerging threats and "
                                    "vulnerabilities. Security & Risk management activities are tracked in the "
                                    "CyberSecure 360 GRC Management platform.",
        "roles_resp_admin_intro": "The [OSC] CUI system environment is serviced by InfoDefense (Managed System "
                                  "Security Provider - MSSP) and [Vendor] (External Service Provider - ESP) with "
                                  "roles, responsibilities, and functions outlined in the SSP (e.g., AC.L2-3.1.2, "
                                  "AC.L2-3.1.4, AC.L2-3.1.5) and outlined by assessment objective in the "
                                  "customer’s Shared Responsibility Matrix (SRM):",
        "roles_resp_srm_items": "[OSC] – Shared Responsibility Matrix – CyberSecure 360 (hyperlink)\n"
                                "Other ESP/CSP Customer Responsibility matrices",
        "roles_resp_admin_closing": "In accordance with NIST SP 800-171, the [OSC] implements compensating "
                                    "controls to reduce the risk of unauthorized or unintentional actions. These "
                                    "controls include management oversight, audit log generation and review, "
                                    "role-based access control and review, user assigned permissions abiding to "
                                    "the rule of least privilege, and periodic review of privileged access.\n\n"
                                    "Separation of duties and role-based (least privilege) access is established "
                                    "and maintained by [OSC] in accordance with the documented organizational "
                                    "policy. Account creation, modification and removal are executed by [OSC] in "
                                    "accordance with the Change Management program as defined in documented "
                                    "organizational policy.",
        "security_awareness_admin_intro": "Through InfoDefense, all [OSC] employees are registered in the "
                                          "KnowBe4 training platform and are required to accomplish annual or "
                                          "remedial security awareness training. This training is comprised of:",
        "security_awareness_admin_items": "Annual Security Awareness Training\n"
                                          "Annual Insider Threat Awareness\n"
                                          "Annual Privileged User Training\n"
                                          "Monthly Simulated Phish Tests\n"
                                          "DoD CUI Handling Training",
        "vulnerability_mgmt_desc": "InfoDefense manages the Vulnerability Management Program as a part of "
                                   "[OSC]’s CyberSecure 360 subscription. InfoDefense implements Greenbone "
                                   "Security Assistant, Microsoft Defender, and Wazuh SIEM integration services "
                                   "and applications to monitor and report vulnerabilities in the system. "
                                   "InfoDefense provides periodic vulnerability reports to the [OSC] stakeholders "
                                   "and system owners for timely resolution. Vulnerability reporting includes "
                                   "detailed information on events, incidents, or vulnerabilities categorized as "
                                   "Critical, High, Medium, or Low severity.",

        "tech_controls_intro": "The following controls are automated in nature, providing logical protection "
                               "of CUI material.",
        "audit_logging_desc": "InfoDefense manages the Audit and Logging services as a part of [OSC]’s "
                              "CyberSecure 360 subscription, and includes audit logging configurations, "
                              "centralized log management, security monitoring services, and security event "
                              "correlation services. Audit and logging services are configured to alert on a "
                              "variety of conditions. Upon discovery and review, InfoDefense Security Analysts "
                              "alert [OSC] stakeholders and system owner of the logged event for remediation. "
                              "Security event logs are retained in accordance with the organization’s "
                              "Records Retention Policy.",
        "config_mgmt_intro": "[OSC] Endpoint configurations are managed and sustained in the organizational "
                             "Microsoft 365 GCC(-High) tenant. System updates are deployed utilizing M365 "
                             "Endpoint Manager. Servers, network appliances and other devices are maintained "
                             "manually, and updates are tracked via the iTop ticketing system (Change Review). "
                             "System baselines and their configuration are maintained in accordance with "
                             "established organizational documentation:",
        "config_mgmt_items": "FortiGate Security Standard\nWindows 11 Security Standard\nWindows Intune Security Standard",
        "encryption_intro": "[OSC] stores all CUI within specific restricted network locations. These "
                            "locations are encrypted utilizing FIPS 140-validated encryption modules "
                            "implemented via the Microsoft 365 GCC-High tenant.\n\n"
                            "[OSC] utilizes BitLocker, managed by Microsoft 365 Endpoint Manager for endpoint "
                            "disk encryption. Microsoft SAML login with Microsoft Entra ID as the IDP is "
                            "utilized to provide multi-factor authentication. The [OSC] backup solution of CUI "
                            "material includes Microsoft SharePoint residing on the organization’s "
                            "Microsoft 365 GCC-High tenant. [OSC] protects CUI material stored in organizational "
                            "systems implementing the native operating system encryption mechanisms:",
        "encryption_items": "Microsoft Encrypted File System (EFS) (folder, file-level)\n"
                            "Microsoft BitLocker (full-disk encryption)\n"
                            "Encryption keys (operating system)",
        "encryption_closing": "Fortinet FortiGate firewalls are configured to provide IPSEC VPNs to connect to "
                              "remote site. The encryption utilized for these sessions leverages FIPS "
                              "140-validated encryption modules to safeguard the confidentiality of CUI material "
                              "sent over remote sessions.",
        "malware_protection_desc": "[OSC] utilizes Fortinet FortiGate firewalls to provide advanced threat "
                                   "protection from malware and intrusion threats. In addition, [OSC] protects "
                                   "endpoints from Malware and other endpoint specific threats through "
                                   "enterprise-wide deployment of Microsoft Defender which is systemically "
                                   "configured to auto-update and provide real-time protection as well as a "
                                   "weekly scan (full scan) on each endpoint.",
        "media_handling_intro": "[OSC] does/does not process, transmit, or store physical / hardcopy CUI "
                                "material. This material is protected … [define how physical / hardcopy "
                                "material is protected at each site (if any)]. [OSC] employees are trained to "
                                "protect CUI material on logical storage devices:",
        "media_handling_items": "Implementing FIPS validated encryption for Data at Rest\n"
                                "Maintaining positive control and accountability of removable media in-transit\n"
                                "Utilizing only authorized / whitelisted portable storage devices\n"
                                "Sanitizing or destroying media in accordance with established organizational policy",
        "network_security_desc": "The [OSC] Home, Satellite, and Remote offices are segregated from any "
                                 "personal home network and the Internet with a Fortinet FortiGate firewall that "
                                 "provides network intrusion prevention, advanced threat protection, and web "
                                 "content filtering.\n\n"
                                 "Firewalls are sized according to bandwidth and other requirements. Networks "
                                 "are segregated by security risk. For example, IP phone, Internet accessible "
                                 "servers, and guest networks are isolated in a DMZ.\n\n"
                                 "Firewall policies are set to “deny all” network communications by "
                                 "default and only allow protocols and services that are required by specific "
                                 "business processes.\n\n"
                                 "FIPS mode is enabled for all firewalls supporting the operational system "
                                 "environment.",
        "network_security_wireless_desc": "The [OSC] does/does not employ wireless access points, and leverages "
                                          "VPN connectivity into the Home, Satellite, and Remote offices when "
                                          "traveling or for remote users. The VPN is configured as an IPSEC VPN "
                                          "tunnel with the application installed via Intune. Only authorized "
                                          "devices and users can access [OSC]’s VPN and are authenticated "
                                          "via web browser using a SAML login with Microsoft Entra ID as the IDP.",
        "physical_security_tech_desc": "Visitor sign-in and positive identification are mandatory for any "
                                       "visitors to a/an [OSC] site. Visitors are required to be escorted by an "
                                       "employee of [OSC] for the duration of their stay.\n\n"
                                       "[OSC] utilizes [vendor specific Protection Access Control Systems (PACS) "
                                       "and/or Protection Access Devices (PADS)] to integrate PADS at all "
                                       "ingress / egress points safeguarding the systems and physical assets "
                                       "processing, storing, or transmitting CUI material. The PACS [does/does "
                                       "not] centrally manage logs, while the PADS [do/do not] enable "
                                       "device-centric logging to the controlled workspaces. Additionally, "
                                       "[address all PACS and PADS per site and for] all rooms processing, "
                                       "storing, or transmitting CUI material. PACS and PADS logs are retained "
                                       "in accordance with the organization’s Records Retention Policy.",
        "risk_analysis_tech_desc": "InfoDefense performs Security & Risk analysis as a part of [OSC]’s "
                                   "CyberSecure 360 subscription. Security & Risk are continuously evaluated "
                                   "upon a significant change to the operational system environment and/or upon "
                                   "discovery of a cybersecurity incident. Periodically, the [OSC] risk posture "
                                   "is reviewed and adjusted in response to emerging threats and vulnerabilities. "
                                   "Security & Risk management activities are tracked in the CyberSecure 360 GRC "
                                   "Management platform.",
        "roles_resp_tech_desc": "In accordance with NIST SP 800-171, the [OSC] implements compensating "
                                "controls to reduce the risk of unauthorized or unintentional actions. These "
                                "controls include management oversight, audit log generation and review, "
                                "role-based access control and review, user assigned permissions abiding to the "
                                "rule of least privilege, and periodic review of privileged access.\n\n"
                                "Separation of duties and role-based (least privilege) access is established "
                                "and maintained by [OSC] in accordance with the documented organizational "
                                "policy. Account creation, modification and removal are executed by [OSC] in "
                                "accordance with the Change Management program as defined in documented "
                                "organizational policy.",
        "security_awareness_tech_desc": "Through InfoDefense, all [OSC] employees are registered in the "
                                        "KnowBe4 training platform and are required to accomplish annual or "
                                        "remedial security awareness training. Training consists of "
                                        "cybersecurity awareness training, simulated phish tests, insider threat "
                                        "awareness, and privileged user awareness training. Included in the "
                                        "KnowBe4 training platform is reoccurring metric reports, phishing test "
                                        "results, and compliance reporting on annual training completion. "
                                        "Separate from the KnowBe4 training platform, InfoDefense provides [OSC] "
                                        "with the CUI Media Handling Procedure.",
        "system_update_desc": "[OSC] updates and maintains systems to their latest practical revision level. "
                              "Operating system, appliance, and application vendor patch notification sources "
                              "are monitored by InfoDefense to ensure notification of updates and requirements "
                              "for security patches are applied in accordance with the documented organizational "
                              "policy (System Update Security Standard).",

        "controls_status": {},
        "controls_narratives": {},
        "sprs_partial": {},
        "gap_analysis": {},
        "impact_categorization": [
            {"category": "FCI",  "confidentiality": "Low",    "integrity": "Low", "availability": "Low", "authority": "GSA"},
            {"category": "CUI",  "confidentiality": "Medium", "integrity": "Low", "availability": "Low", "authority": "NARA"},
            {"category": "ITAR", "confidentiality": "Medium", "integrity": "Low", "availability": "Low", "authority": "US Dept. of State"},
        ],
        "images": {
            "logo": None,                # single image — one org logo on the cover page
            "network_diagram": [],       # lists — a site can have more than one diagram
            "physical_diagram": [],
            "cui_flow_diagram": [],
        },
        "version_history": [
            {"version": "1.0", "date": datetime.now().strftime("%m/%d/%Y"), "author": "", "description": "Initial SSP"}
        ],
    }

# ─────────────────────────────────────────────────────────────
# Routes – Main Pages
# ─────────────────────────────────────────────────────────────
@app.route('/')
def index():
    ssp = load_ssp()
    assets = load_assets()
    return render_template('index.html', ssp=ssp,
                           asset_count=len(assets.get('assets', [])),
                           sprs=compute_sprs(ssp))

@app.route('/system-info', methods=['GET', 'POST'])
@app.route('/system-information', methods=['GET', 'POST'])
def system_info():
    ssp = load_ssp()
    if request.method == 'POST':
        data = request.form.to_dict()
        # Flat fields
        for field in ['org_name','org_abbr','mssp_name','prepared_by','system_name',
                      'system_type','system_categorization','system_uid','version',
                      'last_updated','end_users','privileged_users','cui_categories',
                      'nara_cui_intro']:
            if field in data:
                ssp[field] = data[field]
        # Locations array — sent as multiple location[] values
        locations_raw = request.form.getlist('location[]')
        ssp['locations'] = [loc.strip() for loc in locations_raw if loc.strip()]
        # NARA CUI Categories — sent as multiple nara_cui_category[] values
        nara_cats_raw = request.form.getlist('nara_cui_category[]')
        ssp['nara_cui_categories'] = [c.strip() for c in nara_cats_raw if c.strip()]
        # Keep legacy fields populated from first two locations for backwards compat
        ssp['main_office_site'] = ssp['locations'][0] if len(ssp['locations']) > 0 else ''
        ssp['home_office_site'] = ssp['locations'][1] if len(ssp['locations']) > 1 else ''
        # Nested person objects
        for key in ['info_owner', 'system_owner', 'sso']:
            ssp[key] = {
                'name': data.get(f'{key}_name', ''),
                'title': data.get(f'{key}_title', ''),
                'email': data.get(f'{key}_email', ''),
                'phone': data.get(f'{key}_phone', ''),
            }
        ssp['last_updated'] = datetime.now().strftime("%m.%d.%Y")
        save_ssp(ssp)
        return redirect(url_for('system_info') + '?saved=1')
    # Migrate legacy location fields to locations array if needed
    if not ssp.get('locations'):
        locs = []
        if ssp.get('main_office_site'): locs.append(ssp['main_office_site'])
        if ssp.get('home_office_site'): locs.append(ssp['home_office_site'])
        ssp['locations'] = locs
    return render_template('system_info.html', ssp=ssp,
                           authorized_users_count=len(load_assets().get('authorized_users', [])))

@app.route('/environment', methods=['GET', 'POST'])
def environment():
    ssp = load_ssp()
    if request.method == 'POST':
        # Browsers normalize textarea line endings to CRLF on submit; normalize
        # back to \n so multi-paragraph defaults (joined with plain \n\n) don't
        # pick up extra blank lines each time the form is saved.
        def _nl(s):
            return (s or '').replace('\r\n', '\n')

        for field in ['system_environment', 'cui_data_flow', 'physical_security_desc', 'network_desc',
                      'server_desc', 'workstation_desc', 'crma_mobile_desc', 'specialized_assets_desc',
                      'spa_wireless_desc', 'spa_usb_desc', 'cloud_apps_desc', 'out_of_scope_assets_desc',
                      'main_office_site', 'home_office_site']:
            ssp[field] = _nl(request.form.get(field, ''))

        # Per-location narratives — one textarea per current location, submitted in
        # the same order the page rendered ssp['locations'], so they zip back up
        # by position rather than needing location names as form-field names.
        locations = ssp.get('locations', [])
        phys_vals = [_nl(v) for v in request.form.getlist('phys_sec_narrative[]')]
        ssp['physical_security_by_location'] = {
            loc: phys_vals[i] for i, loc in enumerate(locations) if i < len(phys_vals) and phys_vals[i].strip()
        }
        spa_vals = [_nl(v) for v in request.form.getlist('spa_network_narrative[]')]
        ssp['spa_network_by_location'] = {
            loc: spa_vals[i] for i, loc in enumerate(locations) if i < len(spa_vals) and spa_vals[i].strip()
        }

        # Server types — simple add/remove list, same pattern as NARA CUI Categories
        server_types_raw = request.form.getlist('server_type[]')
        ssp['server_types'] = [s.strip() for s in server_types_raw if s.strip()]

        save_ssp(ssp)
        return redirect(url_for('environment') + '?saved=1')

    locations = ssp.get('locations', [])
    phys_sec_by_loc = {loc: ssp.get('physical_security_by_location', {}).get(loc) or default_physical_security_text(loc)
                       for loc in locations}
    spa_net_by_loc = {loc: ssp.get('spa_network_by_location', {}).get(loc) or default_spa_network_text(loc, locations)
                      for loc in locations}
    return render_template('environment.html', ssp=ssp, phys_sec_by_loc=phys_sec_by_loc, spa_net_by_loc=spa_net_by_loc)

# Free-text narrative fields on the Controls page (Administrative Controls / Technical
# Controls). Everything here is saved/loaded verbatim; CRLF is normalized on save so
# multi-paragraph defaults (joined with plain \n\n) don't pick up extra blank lines.
CONTROLS_TEXT_FIELDS = [
    'controls_intro',
    'admin_controls_intro', 'policies_intro', 'policies_list', 'standards_list', 'procedures_list',
    'access_control_desc', 'asset_mgmt_desc', 'audit_assessment_desc', 'authentication_desc',
    'compliance_mgmt_desc', 'incident_response_desc', 'personnel_security_desc',
    'risk_analysis_admin_desc', 'roles_resp_admin_intro', 'roles_resp_srm_items',
    'roles_resp_admin_closing', 'security_awareness_admin_intro', 'security_awareness_admin_items',
    'vulnerability_mgmt_desc',
    'tech_controls_intro', 'audit_logging_desc', 'config_mgmt_intro', 'config_mgmt_items',
    'encryption_intro', 'encryption_items', 'encryption_closing', 'malware_protection_desc',
    'media_handling_intro', 'media_handling_items', 'network_security_desc',
    'network_security_wireless_desc', 'physical_security_tech_desc', 'risk_analysis_tech_desc',
    'roles_resp_tech_desc', 'security_awareness_tech_desc', 'system_update_desc',
]

@app.route('/controls-section', methods=['GET', 'POST'])
def controls_section():
    ssp = load_ssp()
    if request.method == 'POST':
        def _nl(s):
            return (s or '').replace('\r\n', '\n')
        for field in CONTROLS_TEXT_FIELDS:
            ssp[field] = _nl(request.form.get(field, ''))
        save_ssp(ssp)
        return redirect(url_for('controls_section') + '?saved=1')
    return render_template('controls_section.html', ssp=ssp)

def default_physical_security_text(location):
    return (f"{location}:  All IT systems, hardware, and physical files within the {location} are … "
            "[define the physical access controls (PACS) and physical access devices (PADS) employed at "
            "the location].")

def default_spa_network_text(location, locations):
    """Main site (first location) gets the fuller Fortinet narrative; every other
    site gets the shorter 'connects back to main site' variant, matching the
    source document's Main/Satellite/Remote structure regardless of how many
    locations are actually defined."""
    if locations and location == locations[0]:
        return (f"The [OSC] {location} utilizes IPSEC tunneling via Fortinet firewall technology.\n\n"
                "The FortiGate firewall provides Network Intrusion Prevention, Malware Protection, Anti-Virus "
                "protection and Web Content Filtering through FortiGuard services in addition to "
                "application-aware stateful inspection firewall.")
    main = locations[0] if locations else 'Main Office'
    return f"The [OSC] connects to the {main} via IPSEC tunnels utilizing Fortinet firewall technology."

IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp'}
MULTI_IMAGE_KEYS = ('network_diagram', 'physical_diagram', 'cui_flow_diagram')

@app.route('/images', methods=['GET', 'POST'])
def images():
    ssp = load_ssp()
    if request.method == 'POST':
        # Logo — single image, replaces whatever was there before.
        file = request.files.get('logo')
        if file and file.filename:
            ext = Path(file.filename).suffix.lower()
            if ext in IMAGE_EXTS:
                fname = f"logo_{uuid.uuid4().hex[:8]}{ext}"
                file.save(UPLOADS_DIR / fname)
                old = ssp['images'].get('logo')
                ssp['images']['logo'] = fname
                if old:
                    (UPLOADS_DIR / old).unlink(missing_ok=True)

        # Network / physical / CUI-flow diagrams — each input accepts multiple
        # files at once (browser multi-select), and every save *adds* to the
        # list already on file rather than replacing it, so diagrams can be
        # built up across more than one upload (e.g. one per site).
        for img_key in MULTI_IMAGE_KEYS:
            files = request.files.getlist(img_key)
            ssp['images'].setdefault(img_key, [])
            for file in files:
                if not file or not file.filename:
                    continue
                ext = Path(file.filename).suffix.lower()
                if ext not in IMAGE_EXTS:
                    continue
                fname = f"{img_key}_{uuid.uuid4().hex[:8]}{ext}"
                file.save(UPLOADS_DIR / fname)
                ssp['images'][img_key].append(fname)

        save_ssp(ssp)
        return redirect(url_for('images') + '?saved=1')
    return render_template('images.html', ssp=ssp)

@app.route('/api/images/<img_key>/<filename>', methods=['DELETE'])
def api_delete_image(img_key, filename):
    """Remove one uploaded image. For the multi-image diagram slots this
    removes just that file from the list; the logo slot only ever holds one."""
    if img_key not in MULTI_IMAGE_KEYS and img_key != 'logo':
        return jsonify({"error": "Unknown image type"}), 400
    ssp = load_ssp()
    if img_key == 'logo':
        if ssp['images'].get('logo') != filename:
            return jsonify({"error": "Not found"}), 404
        ssp['images']['logo'] = None
    else:
        current = ssp['images'].get(img_key, [])
        if filename not in current:
            return jsonify({"error": "Not found"}), 404
        ssp['images'][img_key] = [f for f in current if f != filename]
    save_ssp(ssp)
    # filename is validated against the ssp's own recorded image list above,
    # so it can only ever be a name this app itself generated and saved.
    (UPLOADS_DIR / filename).unlink(missing_ok=True)
    return jsonify({"success": True})

@app.route('/assets', methods=['GET', 'POST'])
def assets():
    data = load_assets()
    ssp = load_ssp()
    locations = ssp.get('locations', [])
    if not locations:
        if ssp.get('main_office_site'): locations.append(ssp['main_office_site'])
        if ssp.get('home_office_site'): locations.append(ssp['home_office_site'])
    return render_template('assets.html',
                           assets=data.get('assets', []),
                           physical=data.get('physical_security', []),
                           authorized_users=data.get('authorized_users', []),
                           server_apps=data.get('server_apps', []),
                           mobile_whitelist=data.get('mobile_whitelist', []),
                           system_processes=data.get('system_processes', []),
                           ports_protocols=data.get('ports_protocols', []),
                           blacklist_categories=data.get('blacklist_categories', []),
                           blacklist_tech_impl=data.get('blacklist_tech_impl', ''),
                           blacklist_notice=data.get('blacklist_notice', ''),
                           categories=ASSET_CATEGORIES,
                           device_types=DEVICE_TYPES,
                           physical_types=PHYSICAL_DEVICE_TYPES,
                           locations=locations,
                           personnel_categories=PERSONNEL_CATEGORIES,
                           risk_levels=RISK_LEVELS)

@app.route('/controls', methods=['GET', 'POST'])
def controls():
    ssp = load_ssp()
    if request.method == 'POST':
        ssp['controls_status'] = {}
        ssp.setdefault('controls_narratives', {})
        for key, val in request.form.items():
            if key.startswith('ctrl_status_'):
                ssp['controls_status'][key[12:]] = val
            elif key.startswith('ctrl_narr_'):
                # key format: ctrl_narr_AC.L2-3.1.1_a
                parts = key[10:].rsplit('_', 1)
                if len(parts) == 2:
                    ctrl_id, letter = parts
                    ssp['controls_narratives'].setdefault(ctrl_id, {})[letter] = val
        save_ssp(ssp)
        return redirect(url_for('controls') + '?saved=1')
    return render_template('controls.html', ssp=ssp, controls=CMMC_CONTROLS,
                           ref_map=ASSET_REFERENCE_MAP, default_narratives=CONTROLS_NARRATIVES)

@app.route('/gap-analysis', methods=['GET', 'POST'])
def gap_analysis_page():
    ssp = load_ssp()
    if request.method == 'POST':
        ssp.setdefault('gap_analysis', {})
        for ctrl in CMMC_CONTROLS:
            for letter, _ in ctrl['objectives']:
                prefix = f"{ctrl['id']}_{letter}"
                entry = {
                    'status': request.form.get(f'gap_status_{prefix}', ''),
                    'notes': request.form.get(f'gap_notes_{prefix}', ''),
                    'gap_deficiency': request.form.get(f'gap_deficiency_{prefix}', ''),
                    'gap_action': request.form.get(f'gap_action_{prefix}', ''),
                    'owner': request.form.get(f'gap_owner_{prefix}', ''),
                    'target_date': request.form.get(f'gap_date_{prefix}', ''),
                }
                if any(v.strip() for v in entry.values()):
                    ssp['gap_analysis'].setdefault(ctrl['id'], {})[letter] = entry
                elif ctrl['id'] in ssp['gap_analysis']:
                    # Fully blank row — don't let ssp_data.json accumulate 320 empty entries.
                    ssp['gap_analysis'][ctrl['id']].pop(letter, None)
                    if not ssp['gap_analysis'][ctrl['id']]:
                        del ssp['gap_analysis'][ctrl['id']]
        save_ssp(ssp)
        poam_data = load_poam()
        synced = sync_gap_deficiencies_to_poam(ssp, poam_data)
        if synced:
            save_poam(poam_data)
        return redirect(url_for('gap_analysis_page') + f'?saved=1&poam_synced={synced}')
    return render_template('gap_analysis.html', ssp=ssp, controls=CMMC_CONTROLS,
                           gap_ref=GAP_QUESTIONS, gap_statuses=GAP_STATUSES,
                           gap_data=ssp.get('gap_analysis', {}),
                           summary=gap_summary(ssp))

@app.route('/sra', methods=['GET', 'POST'])
def sra_page():
    data = load_sra()
    if request.method == 'POST':
        data.setdefault('items', {})
        for it in SRA_ITEMS:
            iid = it['id']
            entry = {
                'observations': request.form.get(f'sra_observations_{iid}', ''),
                'interview': request.form.get(f'sra_interview_{iid}', ''),
                'test': request.form.get(f'sra_test_{iid}', ''),
                'criteria': request.form.get(f'sra_criteria_{iid}', ''),
                'notes': request.form.get(f'sra_notes_{iid}', ''),
            }
            if any(v.strip() for v in entry.values()):
                data['items'][iid] = entry
            else:
                data['items'].pop(iid, None)
        save_sra(data)
        return redirect(url_for('sra_page') + '?saved=1')
    return render_template('sra.html', ssp=load_ssp(), families=sra_grouped(),
                           statuses=SRA_STATUSES, sra_data=data.get('items', {}),
                           summary=sra_summary(data))

@app.route('/generate')
def generate():
    ssp = load_ssp()
    assets = load_assets()
    return render_template('generate.html', ssp=ssp, assets=assets)

# ─────────────────────────────────────────────────────────────
# API Routes – Asset CRUD
# ─────────────────────────────────────────────────────────────
@app.route('/api/locations', methods=['GET'])
def api_get_locations():
    """Return the list of office locations defined in System Information."""
    ssp = load_ssp()
    locations = ssp.get('locations', [])
    # Migrate legacy fields if needed
    if not locations:
        if ssp.get('main_office_site'): locations.append(ssp['main_office_site'])
        if ssp.get('home_office_site'): locations.append(ssp['home_office_site'])
    return jsonify(locations)

@app.route('/api/assets', methods=['GET'])
def api_get_assets():
    return jsonify(load_assets())

@app.route('/api/assets', methods=['POST'])
def api_add_asset():
    data = load_assets()
    asset = request.json
    asset['id'] = str(uuid.uuid4())
    asset['created'] = datetime.now().isoformat()
    data.setdefault('assets', []).append(asset)
    save_assets(data)
    return jsonify({"success": True, "id": asset['id']})

@app.route('/api/assets/<asset_id>', methods=['PUT'])
def api_update_asset(asset_id):
    data = load_assets()
    for i, a in enumerate(data.get('assets', [])):
        if a['id'] == asset_id:
            updated = request.json
            updated['id'] = asset_id
            updated['created'] = a.get('created', '')
            updated['modified'] = datetime.now().isoformat()
            data['assets'][i] = updated
            save_assets(data)
            return jsonify({"success": True})
    return jsonify({"error": "Not found"}), 404

@app.route('/api/assets/<asset_id>', methods=['DELETE'])
def api_delete_asset(asset_id):
    data = load_assets()
    data['assets'] = [a for a in data.get('assets', []) if a.get('id') != asset_id]
    save_assets(data)
    return jsonify({"success": True})

@app.route('/api/physical', methods=['POST'])
def api_add_physical():
    data = load_assets()
    item = request.json
    item['id'] = str(uuid.uuid4())
    data.setdefault('physical_security', []).append(item)
    save_assets(data)
    return jsonify({"success": True, "id": item['id']})

@app.route('/api/physical/<item_id>', methods=['PUT'])
def api_update_physical(item_id):
    data = load_assets()
    for i, p in enumerate(data.get('physical_security', [])):
        if p['id'] == item_id:
            updated = request.json
            updated['id'] = item_id
            data['physical_security'][i] = updated
            save_assets(data)
            return jsonify({"success": True})
    return jsonify({"error": "Not found"}), 404

@app.route('/api/physical/<item_id>', methods=['DELETE'])
def api_delete_physical(item_id):
    data = load_assets()
    data['physical_security'] = [p for p in data.get('physical_security', []) if p.get('id') != item_id]
    save_assets(data)
    return jsonify({"success": True})

# ─────────────────────────────────────────────────────────────
# Generic CRUD for additional inventory sections:
#   authorized_users, server_apps, system_processes, ports_protocols
# ─────────────────────────────────────────────────────────────
VALID_SECTIONS = {'authorized_users', 'server_apps', 'mobile_whitelist', 'system_processes', 'ports_protocols'}

@app.route('/api/section/<section>', methods=['GET'])
def api_get_section(section):
    if section not in VALID_SECTIONS:
        return jsonify({"error": "Invalid section"}), 400
    data = load_assets()
    return jsonify(data.get(section, []))

@app.route('/api/section/<section>', methods=['POST'])
def api_add_section_item(section):
    if section not in VALID_SECTIONS:
        return jsonify({"error": "Invalid section"}), 400
    data = load_assets()
    item = request.json
    item['id'] = str(uuid.uuid4())
    item['created'] = datetime.now().isoformat()
    data.setdefault(section, []).append(item)
    save_assets(data)
    return jsonify({"success": True, "id": item['id']})

@app.route('/api/section/<section>/<item_id>', methods=['PUT'])
def api_update_section_item(section, item_id):
    if section not in VALID_SECTIONS:
        return jsonify({"error": "Invalid section"}), 400
    data = load_assets()
    items = data.get(section, [])
    for i, item in enumerate(items):
        if item.get('id') == item_id:
            updated = request.json
            updated['id'] = item_id
            updated['created'] = item.get('created', '')
            updated['modified'] = datetime.now().isoformat()
            items[i] = updated
            save_assets(data)
            return jsonify({"success": True})
    return jsonify({"error": "Not found"}), 404

@app.route('/api/section/<section>/<item_id>', methods=['DELETE'])
def api_delete_section_item(section, item_id):
    if section not in VALID_SECTIONS:
        return jsonify({"error": "Invalid section"}), 400
    data = load_assets()
    data[section] = [item for item in data.get(section, []) if item.get('id') != item_id]
    save_assets(data)
    return jsonify({"success": True})

# ─────────────────────────────────────────────────────────────
# Blacklist CRUD
# ─────────────────────────────────────────────────────────────

@app.route('/api/blacklist', methods=['GET'])
def api_get_blacklist():
    data = load_assets()
    return jsonify({
        "categories": data.get('blacklist_categories', []),
        "tech_impl": data.get('blacklist_tech_impl', ''),
        "notice": data.get('blacklist_notice', ''),
    })

@app.route('/api/blacklist/category', methods=['POST'])
def api_add_blacklist_category():
    """Add a new category"""
    data = load_assets()
    item = request.json
    item['id'] = str(uuid.uuid4())
    item.setdefault('apps', [])
    data.setdefault('blacklist_categories', []).append(item)
    save_assets(data)
    return jsonify({"success": True, "id": item['id']})

@app.route('/api/blacklist/category/<cat_id>', methods=['PUT'])
def api_update_blacklist_category(cat_id):
    """Update category name/justification"""
    data = load_assets()
    cats = data.get('blacklist_categories', [])
    for cat in cats:
        if cat['id'] == cat_id:
            cat.update({k: v for k, v in request.json.items() if k != 'id'})
            save_assets(data)
            return jsonify({"success": True})
    return jsonify({"error": "Not found"}), 404

@app.route('/api/blacklist/category/<cat_id>', methods=['DELETE'])
def api_delete_blacklist_category(cat_id):
    data = load_assets()
    data['blacklist_categories'] = [c for c in data.get('blacklist_categories', []) if c['id'] != cat_id]
    save_assets(data)
    return jsonify({"success": True})

@app.route('/api/blacklist/category/<cat_id>/app', methods=['POST'])
def api_add_blacklist_app(cat_id):
    """Add an app to a category"""
    data = load_assets()
    app_name = (request.json or {}).get('name', '').strip()
    if not app_name:
        return jsonify({"error": "App name required"}), 400
    for cat in data.get('blacklist_categories', []):
        if cat['id'] == cat_id:
            if app_name not in cat.get('apps', []):
                cat.setdefault('apps', []).append(app_name)
            save_assets(data)
            return jsonify({"success": True})
    return jsonify({"error": "Category not found"}), 404

@app.route('/api/blacklist/category/<cat_id>/app/<path:app_name>', methods=['DELETE'])
def api_delete_blacklist_app(cat_id, app_name):
    """Remove an app from a category"""
    data = load_assets()
    for cat in data.get('blacklist_categories', []):
        if cat['id'] == cat_id:
            cat['apps'] = [a for a in cat.get('apps', []) if a != app_name]
            save_assets(data)
            return jsonify({"success": True})
    return jsonify({"error": "Category not found"}), 404

@app.route('/api/blacklist/settings', methods=['POST'])
def api_update_blacklist_settings():
    """Update Technical Implementation text and Notice text"""
    data = load_assets()
    body = request.json or {}
    if 'tech_impl' in body:
        data['blacklist_tech_impl'] = body['tech_impl']
    if 'notice' in body:
        data['blacklist_notice'] = body['notice']
    save_assets(data)
    return jsonify({"success": True})

# ─────────────────────────────────────────────────────────────
# SPRS partial-credit settings (IA.L2-3.5.3 MFA / SC.L2-3.13.11 FIPS)
# ─────────────────────────────────────────────────────────────
VALID_CTRL_STATUSES = {'Implemented', 'Planned to be Implemented', 'Not Applicable', 'Not Implemented'}

@app.route('/api/control-status', methods=['POST'])
def api_control_status():
    """
    Autosave a single control's implementation status the moment it changes,
    so the dashboard SPRS score is always current without needing 'Save All'.
    body: {"ctrl_id": "...", "status": "Implemented"|...}
    Returns the recomputed SPRS summary.
    """
    body = request.json or {}
    cid = body.get('ctrl_id', '')
    status = body.get('status', '')
    if cid not in {c['id'] for c in CMMC_CONTROLS}:
        return jsonify({"error": "Unknown control"}), 400
    if status not in VALID_CTRL_STATUSES:
        return jsonify({"error": "Invalid status"}), 400
    ssp = load_ssp()
    ssp.setdefault('controls_status', {})[cid] = status
    save_ssp(ssp)
    return jsonify({"success": True, "sprs": compute_sprs(ssp)})

@app.route('/api/control-status-bulk', methods=['POST'])
def api_control_status_bulk():
    """Persist many statuses in one call (used by the Mark All buttons)."""
    body = request.json or {}
    statuses = body.get('statuses', {})
    valid_ids = {c['id'] for c in CMMC_CONTROLS}
    ssp = load_ssp()
    ssp.setdefault('controls_status', {})
    applied = 0
    for cid, status in statuses.items():
        if cid in valid_ids and status in VALID_CTRL_STATUSES:
            ssp['controls_status'][cid] = status
            applied += 1
    save_ssp(ssp)
    return jsonify({"success": True, "applied": applied, "sprs": compute_sprs(ssp)})

# ─────────────────────────────────────────────────────────────
# Gap Analysis — status autosave (per objective)
# ─────────────────────────────────────────────────────────────
@app.route('/api/gap-status', methods=['POST'])
def api_gap_status():
    """
    Autosave a single objective's Gap Analysis status the moment it changes,
    so the summary tiles are current without needing 'Save All'.
    body: {"ctrl_id": "...", "letter": "a", "status": "Met"|"Partially Met"|"Not Met"|"N/A"|""}
    Returns the recomputed Gap Analysis summary.
    """
    body = request.json or {}
    cid = body.get('ctrl_id', '')
    letter = body.get('letter', '')
    status = body.get('status', '')
    objective_letters = {c['id']: {l for l, _ in c['objectives']} for c in CMMC_CONTROLS}
    if cid not in objective_letters or letter not in objective_letters[cid]:
        return jsonify({"error": "Unknown objective"}), 400
    if status not in GAP_STATUSES and status != '':
        return jsonify({"error": "Invalid status"}), 400
    ssp = load_ssp()
    ssp.setdefault('gap_analysis', {})
    if status:
        ssp['gap_analysis'].setdefault(cid, {}).setdefault(letter, {})['status'] = status
    elif cid in ssp['gap_analysis'] and letter in ssp['gap_analysis'][cid]:
        ssp['gap_analysis'][cid][letter]['status'] = ''
    save_ssp(ssp)
    return jsonify({"success": True, "summary": gap_summary(ssp)})

# ─────────────────────────────────────────────────────────────
# Security & Risk Analysis — status autosave (per item)
# ─────────────────────────────────────────────────────────────
@app.route('/api/sra-status', methods=['POST'])
def api_sra_status():
    """
    Autosave a single item's Criteria status the moment it changes, so the
    summary tiles / domain bars are current without needing 'Save All'.
    body: {"item_id": "S1-R2", "status": "Meets"|"Does Not Meet"|"Not Applicable"|"Not Assessed"|""}
    """
    body = request.json or {}
    item_id = body.get('item_id', '')
    status = body.get('status', '')
    if item_id not in SRA_ITEM_IDS:
        return jsonify({"error": "Unknown item"}), 400
    if status not in SRA_STATUSES and status != '':
        return jsonify({"error": "Invalid status"}), 400
    data = load_sra()
    data.setdefault('items', {})
    if status:
        data['items'].setdefault(item_id, {})['criteria'] = status
    elif item_id in data['items']:
        data['items'][item_id]['criteria'] = ''
    save_sra(data)
    return jsonify({"success": True, "summary": sra_summary(data)})

# ─────────────────────────────────────────────────────────────
# Information Impact Categorization (editable table)
# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# POA&M routes
# ─────────────────────────────────────────────────────────────
@app.route('/poam')
def poam_page():
    ssp = load_ssp()
    data = load_poam()
    items = sorted(data['items'], key=lambda i: (_poam_sort_key(i)))
    return render_template('poam.html',
                           ssp=ssp,
                           items=items,
                           coverage=poam_coverage(items),
                           summary=poam_summary(items, ssp),
                           statuses=POAM_STATUSES,
                           risk_levels=POAM_RISK_LEVELS,
                           identifiers=POAM_IDENTIFIERS,
                           cost_states=POAM_COST_STATES,
                           milestones=POAM_MILESTONES,
                           ctrl_status=ssp.get('controls_status', {}),
                           nist_ref_map={c['id'].split('-', 1)[1]: c['id'] for c in CMMC_CONTROLS},
                           sprs=compute_sprs(ssp))

def _poam_sort_key(item):
    try:
        return (0, int(item.get('poam_id') or 0))
    except (TypeError, ValueError):
        return (1, 0)

@app.route('/api/poam', methods=['GET'])
def api_poam_list():
    ssp = load_ssp()
    items = load_poam()['items']
    return jsonify({'items': items,
                    'coverage': poam_coverage(items),
                    'summary': poam_summary(items, ssp)})

@app.route('/api/poam', methods=['POST'])
def api_poam_add():
    data = load_poam()
    body = request.json or {}
    item = {fld: str(body.get(fld, '')).strip() for fld in POAM_FIELDS}
    if not item['poam_id']:
        existing = [int(i['poam_id']) for i in data['items']
                    if str(i.get('poam_id', '')).isdigit()]
        item['poam_id'] = str(max(existing) + 1 if existing else 1)
    item['id'] = str(uuid.uuid4())
    data['items'].append(item)
    save_poam(data)
    ssp = load_ssp()
    return jsonify({'success': True, 'id': item['id'],
                    'coverage': poam_coverage(data['items']),
                    'summary': poam_summary(data['items'], ssp)})

@app.route('/api/poam/<item_id>', methods=['PUT'])
def api_poam_update(item_id):
    data = load_poam()
    body = request.json or {}
    for it in data['items']:
        if it['id'] == item_id:
            for fld in POAM_FIELDS:
                if fld in body:
                    it[fld] = str(body[fld]).strip()
            save_poam(data)
            ssp = load_ssp()
            return jsonify({'success': True,
                            'coverage': poam_coverage(data['items']),
                            'summary': poam_summary(data['items'], ssp)})
    return jsonify({'error': 'Not found'}), 404

@app.route('/api/poam/<item_id>', methods=['DELETE'])
def api_poam_delete(item_id):
    data = load_poam()
    before = len(data['items'])
    data['items'] = [i for i in data['items'] if i['id'] != item_id]
    if len(data['items']) == before:
        return jsonify({'error': 'Not found'}), 404
    save_poam(data)
    ssp = load_ssp()
    return jsonify({'success': True,
                    'coverage': poam_coverage(data['items']),
                    'summary': poam_summary(data['items'], ssp)})

@app.route('/api/poam/reset', methods=['POST'])
def api_poam_reset():
    """Restore the full default POA&M template (all 110 requirements)."""
    if POAM_FILE.exists():
        POAM_FILE.unlink()
    data = load_poam()
    ssp = load_ssp()
    return jsonify({'success': True, 'count': len(data['items']),
                    'coverage': poam_coverage(data['items']),
                    'summary': poam_summary(data['items'], ssp)})

@app.route('/api/poam/sync-status', methods=['POST'])
def api_poam_sync_status():
    """
    Mark POA&M items Completed when every NIST requirement they cover is
    Implemented or N/A in the controls page. Items with any Not Met
    requirement are left open.
    """
    ssp = load_ssp()
    data = load_poam()
    changed = 0
    for it in data['items']:
        cs = poam_control_status(it, ssp)
        if cs['met'] + cs['not_met'] == 0:
            continue
        if cs['not_met'] == 0 and it.get('status') not in ('Completed', 'Risk Accepted'):
            it['status'] = 'Completed'
            if not it.get('date_complete'):
                it['date_complete'] = datetime.now().strftime('%m/%d/%Y')
            changed += 1
        elif cs['not_met'] > 0 and it.get('status') == 'Completed':
            it['status'] = 'In Progress'
            changed += 1
    save_poam(data)
    return jsonify({'success': True, 'changed': changed,
                    'coverage': poam_coverage(data['items']),
                    'summary': poam_summary(data['items'], ssp)})

@app.route('/export/poam')
def export_poam():
    """Export the POA&M in the standard OSC workbook format."""
    ssp = load_ssp()
    data = load_poam()
    items = sorted(data['items'], key=_poam_sort_key)
    wb = build_poam_xlsx(items, ssp)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    org = ssp.get('org_abbr', 'Customer') or 'Customer'
    return send_file(buf, download_name=f"{org}_CMMC_L2_POAM.xlsx",
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def build_poam_xlsx(items, ssp):
    """Build the POA&M workbook matching the standard 17-column format."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "POA&M"
    ws.sheet_view.showGridLines = False
    org = ssp.get('org_name', 'Organization') or 'Organization'

    headers = ['POA&M\nID #', 'CMMC\nRef.', 'NIST\nRef.', 'Owner', 'Deficiency\nDetails',
               'Deficiency\nDate', 'Deficiency Identifier', 'Remediation\n(Security Controls)',
               'Estimated\nCost', 'Date\nComplete', 'Milestone', 'TTP\nIdentifier',
               'Risk\nLikelihood', 'Risk\nImpact', 'Risk\nLevel', 'POA&M ID Status',
               'Milestones\nNotes']
    widths = [8, 14, 12, 18, 40, 12, 20, 34, 11, 11, 22, 30, 12, 11, 10, 16, 34]

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    t = ws.cell(1, 1)
    t.value = f"Plan of Action & Milestones (POA&M) – {org}"
    t.font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    t.fill = PatternFill('solid', fgColor='1F3864')
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    thin = Side(style='thin', color='BFBFBF')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.row_dimensions[2].height = 42
    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(2, c)
        cell.value = h
        cell.font = Font(bold=True, size=9, color='FFFFFF', name='Arial')
        cell.fill = PatternFill('solid', fgColor='1F3864')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(c)].width = w

    risk_fill = {'High': 'FFC7CE', 'Med': 'FFEB9C', 'Low': 'C6EFCE'}
    status_fill = {'Completed': 'C6EFCE', 'In Progress': 'FFEB9C',
                   'Not Started': 'F2F2F2', 'Risk Accepted': 'BDD7EE', 'Deferred': 'E8E1F0'}

    for r, it in enumerate(items, 3):
        vals = [it.get(f, '') for f in POAM_FIELDS]
        ws.row_dimensions[r].height = 32
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c)
            cell.value = sanitize_text(str(v)) if v else ''
            cell.font = Font(size=9, name='Arial', bold=(c == 1))
            cell.alignment = Alignment(vertical='top', wrap_text=True,
                                       horizontal='center' if c in (1, 9, 13, 14, 15, 16) else 'left')
            cell.border = border
            fld = POAM_FIELDS[c - 1]
            if fld in ('risk_likelihood', 'risk_impact', 'risk_level') and v in risk_fill:
                cell.fill = PatternFill('solid', fgColor=risk_fill[v])
            elif fld == 'status' and v in status_fill:
                cell.fill = PatternFill('solid', fgColor=status_fill[v])
            elif r % 2 == 0:
                cell.fill = PatternFill('solid', fgColor='F7F9FC')

    ws.freeze_panes = 'A3'
    if items:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}2"

    # Coverage sheet — shows every 800-171 requirement and its POA&M item
    ws2 = wb.create_sheet("NIST 800-171 Coverage")
    ws2.sheet_view.showGridLines = False
    cov = poam_coverage(items)
    ws2.merge_cells('A1:E1')
    t2 = ws2['A1']
    t2.value = f"NIST SP 800-171 R2 Coverage — {cov['covered']} of {cov['total_reqs']} requirements mapped"
    t2.font = Font(bold=True, size=12, color='FFFFFF', name='Arial')
    t2.fill = PatternFill('solid', fgColor='1F3864')
    t2.alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[1].height = 26

    cov_headers = ['NIST Ref.', 'CMMC L2 Requirement', 'Requirement Name', 'POA&M ID(s)', 'Control Status']
    for c, (h, w) in enumerate(zip(cov_headers, [12, 18, 46, 14, 18]), 1):
        cell = ws2.cell(2, c)
        cell.value = h
        cell.font = Font(bold=True, size=9, color='FFFFFF', name='Arial')
        cell.fill = PatternFill('solid', fgColor='2F5496')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
        ws2.column_dimensions[get_column_letter(c)].width = w

    statuses = ssp.get('controls_status', {})
    r = 3
    for ctrl in CMMC_CONTROLS:
        ref = ctrl['id'].split('-', 1)[1]
        ids = cov['map'].get(ref, [])
        cstatus = statuses.get(ctrl['id'], 'Not Implemented')
        row_vals = [ref, ctrl['id'], ctrl['name'],
                    ', '.join(str(i) for i in ids) if ids else 'NOT MAPPED', cstatus]
        for c, v in enumerate(row_vals, 1):
            cell = ws2.cell(r, c)
            cell.value = v
            cell.font = Font(size=9, name='Arial')
            cell.alignment = Alignment(vertical='center', wrap_text=(c == 3))
            cell.border = border
            if c == 4 and not ids:
                cell.fill = PatternFill('solid', fgColor='FFC7CE')
                cell.font = Font(size=9, name='Arial', bold=True, color='9C0006')
            elif c == 5:
                fillc = 'C6EFCE' if cstatus in ('Implemented', 'Not Applicable') else (
                    'FFEB9C' if cstatus == 'Planned to be Implemented' else 'FFC7CE')
                cell.fill = PatternFill('solid', fgColor=fillc)
            elif r % 2 == 0:
                cell.fill = PatternFill('solid', fgColor='F7F9FC')
        r += 1
    ws2.freeze_panes = 'A3'
    return wb

# ─────────────────────────────────────────────────────────────
# Export: Gap Analysis XLSX
# ─────────────────────────────────────────────────────────────
GAP_DOMAIN_COLORS = {   # matches the client-supplied source workbook's row shading
    'AC': 'DDEBF7', 'AT': 'FCE4D6', 'AU': 'E2EFDA', 'CM': 'FFF2CC', 'IA': 'D9D2E9',
    'IR': 'F4CCCC', 'MA': 'D9EAD3', 'MP': 'FCE5CD', 'PS': 'CFE2F3', 'PE': 'EAD1DC',
    'RA': 'D0E0E3', 'CA': 'F9CB9C', 'SC': 'C9DAF8', 'SI': 'D9D9D9',
}

@app.route('/export/gap')
def export_gap():
    """Export the Gap Analysis in the same 15-column layout as the source workbook."""
    ssp = load_ssp()
    wb = build_gap_xlsx(ssp)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    org = ssp.get('org_abbr', 'Customer') or 'Customer'
    return send_file(buf, download_name=f"{org}_CMMC_L2_Gap_Analysis.xlsx",
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def build_gap_xlsx(ssp):
    """Build the Gap Analysis workbook: one row per assessment objective (320 rows)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gap Analysis"
    ws.sheet_view.showGridLines = False
    gap_data = ssp.get('gap_analysis', {})

    headers = ['#', 'Domain', 'Domain Name', 'Control ID', 'Control Title',
               'Assessment Objective', 'Objective (Reference)', 'Applies To (System Components)',
               'Gap Analysis Question', 'Suggested Evidence / Source',
               'Status (Met / Partially Met / Not Met / N-A)', 'Evidence Provided / Notes',
               'Gap / Remediation Action', 'Owner', 'Target Date']
    widths = [6, 9, 22, 12, 26, 16, 40, 26, 50, 44, 16, 34, 34, 16, 14]

    thin = Side(style='thin', color='BFBFBF')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(1, c)
        cell.value = h
        cell.font = Font(bold=True, size=9, color='FFFFFF', name='Arial')
        cell.fill = PatternFill('solid', fgColor='1F3864')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 34

    import re as _re
    status_fill = {'Met': 'C6EFCE', 'Partially Met': 'FFEB9C', 'Not Met': 'FFC7CE', 'N/A': 'D9D9D9'}
    r = 2
    n = 0
    for ctrl in CMMC_CONTROLS:
        domain_code = ctrl['id'].split('.', 1)[0]
        ref = ctrl['id'].split('-', 1)[1]
        gref = GAP_QUESTIONS.get(ctrl['id'], {})
        for letter, obj_text in ctrl['objectives']:
            n += 1
            entry = (gap_data.get(ctrl['id']) or {}).get(letter) or {}
            status = entry.get('status', '')
            row_vals = [
                n, domain_code, ctrl['domain'].split(' ', 1)[1] if ' ' in ctrl['domain'] else ctrl['domain'],
                ref, ctrl['name'], f"{ref}[{letter}]", _re.sub(r';?\s*and\s*$', '', obj_text).strip(),
                gref.get('applies_to', ''), gref.get('questions', {}).get(letter, ''),
                gref.get('evidence_source', ''), status,
                sanitize_text(entry.get('notes', '')), sanitize_text(entry.get('gap_action', '')),
                sanitize_text(entry.get('owner', '')), sanitize_text(entry.get('target_date', '')),
            ]
            dom_color = GAP_DOMAIN_COLORS.get(domain_code, 'F2F2F2')
            for c, v in enumerate(row_vals, 1):
                cell = ws.cell(r, c)
                cell.value = v
                cell.font = Font(size=9, name='Arial', bold=(c in (1, 4)))
                cell.alignment = Alignment(vertical='top', wrap_text=(c not in (1, 4, 11, 14, 15)),
                                           horizontal='center' if c in (1, 4, 6, 11, 15) else 'left')
                cell.border = border
                if c == 2:
                    cell.fill = PatternFill('solid', fgColor=dom_color)
                    cell.alignment = Alignment(horizontal='center', vertical='top')
                elif c == 11 and status in status_fill:
                    cell.fill = PatternFill('solid', fgColor=status_fill[status])
            ws.row_dimensions[r].height = 28
            r += 1

    last_row = r - 1
    if last_row >= 2:
        dv = DataValidation(type='list', formula1='"Met,Partially Met,Not Met,N/A"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f'K2:K{last_row}')
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row}"

    # Summary sheet — live status/domain rollup
    ws2 = wb.create_sheet("Summary")
    ws2.sheet_view.showGridLines = False
    summ = gap_summary(ssp)
    org = ssp.get('org_name', 'Organization') or 'Organization'
    ws2.merge_cells('A1:E1')
    t = ws2['A1']
    t.value = f"CMMC L2 Gap Analysis Summary – {org}"
    t.font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    t.fill = PatternFill('solid', fgColor='1F3864')
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[1].height = 30

    row = 3
    for lbl, val in [('Total Objectives', summ['total']), ('Answered', summ['answered']),
                      ('Unanswered', summ['unanswered']), ('% Answered', f"{summ['pct_answered']}%")]:
        ws2.cell(row, 1).value = lbl
        ws2.cell(row, 1).font = Font(bold=True, size=10, name='Arial')
        ws2.cell(row, 2).value = val
        ws2.cell(row, 2).font = Font(size=10, name='Arial')
        row += 1
    row += 1
    ws2.cell(row, 1).value = 'By Status'
    ws2.cell(row, 1).font = Font(bold=True, size=11, name='Arial', color='1F3864')
    row += 1
    for s in GAP_STATUSES:
        ws2.cell(row, 1).value = s
        ws2.cell(row, 1).font = Font(size=10, name='Arial')
        ws2.cell(row, 2).value = summ['counts'][s]
        ws2.cell(row, 2).font = Font(size=10, name='Arial')
        if s in status_fill:
            ws2.cell(row, 2).fill = PatternFill('solid', fgColor=status_fill[s])
        row += 1
    ws2.column_dimensions['A'].width = 22
    ws2.column_dimensions['B'].width = 14
    return wb

# ─────────────────────────────────────────────────────────────
# Export: Security & Risk Analysis (XLSX — matches the source workbook's
# own sheet layout: one tab per control set, plus a Reports tab and a
# Key Remediation Points tab)
# ─────────────────────────────────────────────────────────────
SRA_ITEM_HEADERS = ['Ref. Category', 'Control', 'Objective', 'Observations', 'Interview', 'Examine', 'Test', 'Criteria', 'Notes']
SRA_ITEM_WIDTHS = [30, 46, 40, 24, 24, 32, 20, 17, 32]
SRA_CRITERIA_FILL = {'Meets': 'C6EFCE', 'Does Not Meet': 'FFC7CE', 'Not Applicable': 'D9D9D9', 'Not Assessed': 'F7F9FC'}

@app.route('/export/sra')
def export_sra():
    ssp = load_ssp()
    data = load_sra()
    wb = build_sra_xlsx(ssp, data)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    org = ssp.get('org_abbr', 'Customer') or 'Customer'
    return send_file(buf, download_name=f"{org}_Security_Risk_Analysis.xlsx",
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def _sra_write_item_sheet(wb, sheet_name, sheet_items, sra_data):
    """One tab per control set, matching the source workbook's own column
    layout exactly (Ref. Category / Control / Objective / Observations /
    Interview / Examine / Test / Criteria / Notes), populated with the
    saved assessment data."""
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    thin = Side(style='thin', color='BFBFBF')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c, (h, w) in enumerate(zip(SRA_ITEM_HEADERS, SRA_ITEM_WIDTHS), 1):
        cell = ws.cell(1, c)
        cell.value = h
        cell.font = Font(bold=True, size=10, color='FFFFFF', name='Arial')
        cell.fill = PatternFill('solid', fgColor='1F3864')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 20

    r = 2
    prev_ref = None
    for it in sheet_items:
        entry = sra_data.get('items', {}).get(it['id'], {})
        criteria = entry.get('criteria') or 'Not Assessed'
        ref_cat = f"{it['ref']}: {it['ref_name']}" if it['ref'] != prev_ref else ''
        prev_ref = it['ref']
        row_vals = [ref_cat, it['control_text'], it['objective_text'], entry.get('observations', ''),
                    entry.get('interview', ''), it['examine'], entry.get('test', ''), criteria,
                    entry.get('notes', '')]
        for c, v in enumerate(row_vals, 1):
            cell = ws.cell(r, c)
            cell.value = sanitize_text(v) if v else ''
            cell.font = Font(size=9, name='Arial', bold=(c == 1))
            cell.alignment = Alignment(vertical='top', wrap_text=True)
            cell.border = border
            if c == 8:
                cell.fill = PatternFill('solid', fgColor=SRA_CRITERIA_FILL.get(criteria, 'FFFFFF'))
                cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[r].height = 30
        r += 1

    last_row = r - 1
    if last_row >= 2:
        dv = DataValidation(type='list', formula1='"Not Assessed,Meets,Does Not Meet,Not Applicable"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f'H2:H{last_row}')
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = f"A1:{get_column_letter(len(SRA_ITEM_HEADERS))}{last_row}"
    return ws

def build_sra_xlsx(ssp, sra_data):
    org = ssp.get('org_name', 'Organization') or 'Organization'
    summary = sra_summary(sra_data)
    deficiencies = sra_deficiencies(sra_data)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for sheet in SRA_SHEETS:
        sheet_items = [it for it in SRA_ITEMS if it['sheet'] == sheet]
        _sra_write_item_sheet(wb, sheet, sheet_items, sra_data)

    # ── Reports tab: control-set summary (matches the source workbook's own
    # Reports tab) + a control-family breakdown with a native bar chart ──
    ws = wb.create_sheet('Reports', 0)
    ws.sheet_view.showGridLines = False
    thin = Side(style='thin', color='BFBFBF')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells('A1:F1')
    t = ws['A1']
    t.value = f"Security & Risk Analysis – {org}"
    t.font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    t.fill = PatternFill('solid', fgColor='1F3864')
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws['A3'] = 'Control Assessment Summary'
    ws['A3'].font = Font(bold=True, size=12, name='Arial', color='1F3864')
    headers = ['Control Set', 'Not Assessed', 'Meets', 'Does Not Meet', 'Not Applicable', 'Total']
    for c, h in enumerate(headers, 1):
        cell = ws.cell(4, c)
        cell.value = h
        cell.font = Font(bold=True, size=10, color='FFFFFF', name='Arial')
        cell.fill = PatternFill('solid', fgColor='2F5496')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    r = 5
    for s in summary['by_sheet']:
        vals = [s['label'], s['counts']['Not Assessed'], s['counts']['Meets'],
                s['counts']['Does Not Meet'], s['counts']['Not Applicable'], s['total']]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c)
            cell.value = v
            cell.font = Font(size=10, name='Arial')
            cell.border = border
        r += 1
    ov = summary['overall']
    total_vals = ['Total', ov['counts']['Not Assessed'], ov['counts']['Meets'],
                  ov['counts']['Does Not Meet'], ov['counts']['Not Applicable'], ov['total']]
    for c, v in enumerate(total_vals, 1):
        cell = ws.cell(r, c)
        cell.value = v
        cell.font = Font(bold=True, size=10, name='Arial')
        cell.fill = PatternFill('solid', fgColor='EEF0F3')
        cell.border = border
    r += 3

    fam_hdr_row = r
    ws.cell(r, 1).value = 'Assessment Status by NIST SP 800-53 Control Family'
    ws.cell(r, 1).font = Font(bold=True, size=12, name='Arial', color='1F3864')
    r += 1
    fam_col_hdr_row = r
    fam_headers = ['Family', 'Not Assessed', 'Meets', 'Does Not Meet', 'Not Applicable', 'Total', '% Met']
    for c, h in enumerate(fam_headers, 1):
        cell = ws.cell(r, c)
        cell.value = h
        cell.font = Font(bold=True, size=10, color='FFFFFF', name='Arial')
        cell.fill = PatternFill('solid', fgColor='2F5496')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    r += 1
    fam_data_start = r
    for f in summary['by_family']:
        vals = [f"{f['code']} - {f['name']}", f['counts']['Not Assessed'], f['counts']['Meets'],
                f['counts']['Does Not Meet'], f['counts']['Not Applicable'], f['total'],
                (f['pct_met'] / 100) if f['pct_met'] is not None else None]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c)
            cell.value = v
            cell.font = Font(size=10, name='Arial')
            cell.border = border
            if c == 7 and v is not None:
                cell.number_format = '0%'
        r += 1
    fam_data_end = r - 1

    for col, w in zip('ABCDEFG', [40, 13, 10, 14, 15, 9, 9]):
        ws.column_dimensions[col].width = w

    # Native Excel bar chart — percent met by control family
    chart = BarChart()
    chart.type = 'bar'  # horizontal bars — reads better with 16 category labels
    chart.title = 'Percent Met by Control Family'
    # openpyxl's x_axis is always the category axis and y_axis always the
    # value axis, regardless of bar direction — it's easy to read that
    # backwards for a horizontal chart since visually the value axis runs
    # along the bottom (where an x-axis "should" be).
    chart.x_axis.title = None
    chart.y_axis.title = 'Percent Met'
    chart.y_axis.numFmt = '0%'
    data_ref = Reference(ws, min_col=7, min_row=fam_col_hdr_row, max_row=fam_data_end)
    cats_ref = Reference(ws, min_col=1, min_row=fam_data_start, max_row=fam_data_end)
    chart.add_data(data_ref, titles_from_data=True)
    # chart.set_categories() always writes a numRef, which is wrong for text
    # categories (family names) — build the AxDataSource as a strRef by hand.
    for s in chart.series:
        s.cat = AxDataSource(strRef=StrRef(f=str(cats_ref)))
    chart.height, chart.width = 16, 26
    series = chart.series[0]
    series.graphicalProperties.solidFill = '1F3864'
    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.numFmt = '0%'
    chart.legend = None
    ws.add_chart(chart, f'A{fam_data_end + 3}')

    # ── Key Remediation Points tab ──
    ws3 = wb.create_sheet('Key Remediation Points')
    ws3.sheet_view.showGridLines = False
    ws3.merge_cells('A1:F1')
    t3 = ws3['A1']
    t3.value = f"Key Remediation Points – {len(deficiencies)} item(s) marked Does Not Meet"
    t3.font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    t3.fill = PatternFill('solid', fgColor='B3261E')
    t3.alignment = Alignment(horizontal='center', vertical='center')
    ws3.row_dimensions[1].height = 28

    rem_headers = ['Family', 'Ref.', 'Control / Requirement', 'Observations', 'Recommended Remediation']
    rem_widths = [30, 12, 46, 34, 46]
    for c, (h, w) in enumerate(zip(rem_headers, rem_widths), 1):
        cell = ws3.cell(3, c)
        cell.value = h
        cell.font = Font(bold=True, size=10, color='FFFFFF', name='Arial')
        cell.fill = PatternFill('solid', fgColor='1F3864')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
        ws3.column_dimensions[get_column_letter(c)].width = w
    ws3.row_dimensions[3].height = 20

    r = 4
    if not deficiencies:
        ws3.merge_cells(f'A{r}:E{r}')
        cell = ws3.cell(r, 1)
        cell.value = 'No items are currently marked "Does Not Meet." This section will list remediation points for any control found deficient during the assessment.'
        cell.font = Font(italic=True, size=10, name='Arial', color='5B636E')
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    else:
        for d in deficiencies:
            text = d['objective_text'] or d['control_text']
            remediation = d['notes'] or 'Not yet documented — add remediation guidance in the Notes field on the Security & Risk Analysis page.'
            vals = [f"{d['family']} - {SRA_FAMILIES.get(d['family'], d['family'])}",
                    f"{d['ref']} - {d['ref_name']}", text, d['observations'], remediation]
            for c, v in enumerate(vals, 1):
                cell = ws3.cell(r, c)
                cell.value = sanitize_text(v) if v else ''
                cell.font = Font(size=9, name='Arial', italic=(c == 5 and not d['notes']),
                                 color='5B636E' if (c == 5 and not d['notes']) else '000000')
                cell.alignment = Alignment(vertical='top', wrap_text=True)
                cell.border = border
            ws3.row_dimensions[r].height = 44
            r += 1
        ws3.freeze_panes = 'A4'
        ws3.auto_filter.ref = f"A3:{get_column_letter(len(rem_headers))}{r - 1}"

    wb.active = 0
    return wb

@app.route('/api/impact-categorization', methods=['POST'])
def api_impact_categorization():
    """
    Replace the Information Impact Categorization table.
    body: {"rows": [{"category","confidentiality","integrity","availability","authority"}, ...]}
    Defaults (FCI/CUI/ITAR) are just ordinary rows — removable like any other.
    """
    body = request.json or {}
    rows = body.get('rows', [])
    if not isinstance(rows, list):
        return jsonify({"error": "rows must be a list"}), 400
    clean = []
    for r in rows:
        cat = str(r.get('category', '')).strip()
        if not cat:
            continue  # skip blank rows
        clean.append({
            'category': cat[:80],
            'confidentiality': str(r.get('confidentiality', '')).strip()[:40],
            'integrity': str(r.get('integrity', '')).strip()[:40],
            'availability': str(r.get('availability', '')).strip()[:40],
            'authority': str(r.get('authority', '')).strip()[:120],
        })
    ssp = load_ssp()
    ssp['impact_categorization'] = clean
    save_ssp(ssp)
    return jsonify({"success": True, "count": len(clean)})

@app.route('/api/sprs-partial', methods=['POST'])
def api_sprs_partial():
    """
    Save partial-implementation state for the two 32 CFR 170.24(c)(2)(i)(A)(4)
    partial-credit requirements. body: {"ctrl_id": "...", "mode": "partial"|"none"}
    Returns the recomputed score so the dashboard can update live.
    """
    body = request.json or {}
    cid = body.get('ctrl_id', '')
    mode = body.get('mode', 'none')
    if cid not in SPRS_SPECIAL:
        return jsonify({"error": "Not a partial-credit requirement"}), 400
    if mode not in ('partial', 'none'):
        return jsonify({"error": "mode must be 'partial' or 'none'"}), 400
    ssp = load_ssp()
    ssp.setdefault('sprs_partial', {})[cid] = mode
    save_ssp(ssp)
    return jsonify({"success": True, "sprs": compute_sprs(ssp)})

def _is_example_row(first_cell_text):
    """True for a tracker row that's baked-in template/example content, not
    real customer data — 'Ex.' in the row-number column (Devices/Users/
    Server Apps/Mobile Whitelist/Physical Security), or a '(Example)'
    marker embedded in the cell text itself (System Processes), or an
    '(example)' suffix on the value (Non & Essential PPS)."""
    t = (first_cell_text or '').strip().lower()
    return t == 'ex.' or 'example)' in t

def _col_lookup(colmap, aliases):
    """Resolve a column index from a _hdr_map()'d header row given a list of
    candidate alias strings (already lowercased). Tries, in order: exact
    match, then a header that *starts with* the alias, then a header that
    merely *contains* the alias — so trailing parentheticals like
    'Privileged User \\n(List PIM Roles)' or a shared substring like
    'Essential Program, Function, Port, Protocol, Service' (common to both
    the 'Essential' and 'Non-Essential' sub-table headers) still resolve
    without needing an exact-text alias per sheet variant."""
    for a in aliases:
        if a in colmap:
            return colmap[a]
    for a in aliases:
        for k, idx in colmap.items():
            if k.startswith(a):
                return idx
    for a in aliases:
        for k, idx in colmap.items():
            if a in k:
                return idx
    return None

def _iter_tracker_table_rows(ws, header_markers):
    """Scan a worksheet for one or more sub-tables, each introduced by a
    header row whose first cell case-insensitively matches one of
    `header_markers` (e.g. '#' for the row-numbered sheets, or a literal
    first-column heading for sheets that don't use a row-number column).
    Yields (colmap, row) for every real data row — skipping title/spacer
    rows before the first header, example/template rows (see
    _is_example_row), and rows where every column past the first is empty.
    A row counts as real data even if its own row-number cell is blank —
    tracker rows filled out by hand often leave that column blank on real
    entries while the leftover blank template rows below still carry a
    number.
    """
    markers = {m.lower() for m in header_markers}
    colmap = None
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        first = _cell(row[0])
        if first.lower() in markers:
            colmap = _hdr_map(row)
            continue
        if colmap is None:
            continue
        if _is_example_row(first):
            continue
        other_vals = [row[i] for i in range(1, len(row))]
        if not any(_cell(v) for v in other_vals):
            continue
        yield colmap, row

def _import_generic_table(ws, header_markers, field_map, required_any, skip_row_number_col=True):
    """Extract records from a tracker sheet using _iter_tracker_table_rows().
    field_map: {output_field_name: [lowercased alias strings]} — resolved via
    _col_lookup() per row (so differing column order/casing between a
    sheet's own sub-tables still lines up correctly).
    required_any: record dropped if every one of these output fields is empty.
    Any populated column NOT covered by field_map is preserved in the
    record's '__extras__' string ("Column Header: value; ...") so switching
    layouts between sub-tables never silently drops data.
    Returns a list of dicts, each with the field_map keys plus '__extras__'.
    """
    out = []
    for colmap, row in _iter_tracker_table_rows(ws, header_markers):
        rec = {}
        mapped_idxs = set()
        for fname, aliases in field_map.items():
            idx = _col_lookup(colmap, aliases)
            rec[fname] = _cell(row[idx]) if idx is not None and idx < len(row) else ''
            if idx is not None:
                mapped_idxs.add(idx)
        if not any(rec.get(f) for f in required_any):
            continue
        if skip_row_number_col:
            mapped_idxs.add(0)
        extras = []
        for h, idx in colmap.items():
            if idx in mapped_idxs or idx >= len(row):
                continue
            v = _cell(row[idx])
            if v:
                extras.append(f"{h.replace(chr(10), ' ').strip().title()}: {v}")
        rec['__extras__'] = '; '.join(extras)
        out.append(rec)
    return out

def _dedup_key(d, fields):
    return tuple((d.get(f) or '').strip().lower() for f in fields)

def _append_new(target_list, new_items, key_fields):
    """Append items from new_items that don't already exist in target_list
    (matched by a case-insensitive key over key_fields). Returns count added."""
    existing_keys = {_dedup_key(item, key_fields) for item in target_list}
    added = 0
    for item in new_items:
        key = _dedup_key(item, key_fields)
        if key in existing_keys or not any(key):
            continue
        target_list.append(item)
        existing_keys.add(key)
        added += 1
    return added

def _normalize_blacklist_name(name):
    """Strip a trailing parenthetical qualifier (e.g. '(adjustable as
    needed)' vs. a customer's own '(can be adjusted as needed)') so
    otherwise-identical category names compare equal."""
    return re.sub(r'\s*\([^)]*\)\s*$', '', name or '').strip().lower()

def _find_matching_blacklist_category(existing_categories, name, threshold=0.82):
    """Fuzzy-match a tracker category name against the app's existing
    categories, tolerant of minor wording differences and typos (e.g. a
    tracker's 'Anonomyous Network Applications' vs. the app's default
    'Anonymous Network Applications') so imports update the existing
    category instead of creating a near-duplicate."""
    import difflib
    target = _normalize_blacklist_name(name)
    best, best_ratio = None, 0.0
    for cat in existing_categories:
        ratio = difflib.SequenceMatcher(None, target, _normalize_blacklist_name(cat['name'])).ratio()
        if ratio > best_ratio:
            best, best_ratio = cat, ratio
    return best if best_ratio >= threshold else None

def _parse_blacklist_sheet(ws):
    """Parse the 'Workstation App Blacklist' sheet's bespoke layout:
    category name (col A) + apps across the rest of the row, then a blank
    gap, then '<Category Name> Justification' (col A) + justification text
    (col B) rows, then a blank gap, then standalone 'Technical
    Implementation' and 'Notice' rows. Returns
    (categories, tech_impl, notice) where categories is a list of
    {name, apps, justification} and tech_impl/notice are strings (empty if
    not present).
    """
    categories = []  # [{name, apps: [...]}]
    justifications = {}  # name_lower -> text
    tech_impl = ''
    notice = ''
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        label = _cell(row[0])
        if not label:
            continue
        low = label.lower()
        if low.endswith('justification'):
            name = label[:-len('justification')].strip()
            text = _cell(row[1]) if len(row) > 1 else ''
            if name and text:
                justifications[name.lower()] = text
        elif low == 'technical implementation':
            tech_impl = _cell(row[1]) if len(row) > 1 else ''
        elif low == 'notice':
            notice = _cell(row[1]) if len(row) > 1 else ''
        elif low == 'definitions':
            continue  # trailing glossary block — not category data
        else:
            apps = [_cell(v) for v in row[1:] if _cell(v)]
            if apps:
                categories.append({'name': label.strip(), 'apps': apps})
    for cat in categories:
        cat['justification'] = justifications.get(cat['name'].lower(), '')
    return categories, tech_impl, notice

@app.route('/api/import-xlsx', methods=['POST'])
def api_import_xlsx():
    """Import from a previously exported CMMC L2 Artifacts Tracker Excel
    file across every sheet that holds real per-customer records, routing
    each to its matching Asset Inventory section:
      Authorized Devices        -> IT Assets
      Authorized Users          -> Authorized Users
      Authorized Server Apps    -> Server Apps
      Mobile Whitelist          -> Mobile Whitelist
      Authorized System Processes -> System Processes
      Non & Essential PPS       -> Ports & Protocols
      Physical Security         -> Physical Security
      Workstation App Blacklist -> App Blacklist (updates existing
                                    categories by name, adds new ones, and
                                    updates Technical Implementation / Notice)
    Non-US Access, Contractor Risk Review, Risk Heat Map, and Risk Matrix
    are intentionally NOT imported — they're fixed regulatory/methodology
    reference tables (identical across every tracker export) that are
    already built into this app's Asset Inventory -> Reference tab, not
    per-customer records.

    Each sheet's sub-tables are matched to output fields by header name
    (not fixed column position), since a sheet like Authorized Devices
    holds more than one sub-table with different column layouts. A row
    counts as real data if any column other than the row-number column is
    populated — the row-number ('#') itself is not required, since trackers
    filled out by hand often leave it blank on real entries. Re-running an
    import is safe: records are de-duplicated against what's already saved
    before being appended.
    """
    file = request.files.get('file')
    if not file:
        return jsonify({"error": "No file"}), 400
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file.read()), data_only=True)
        data = load_assets()
        results = {}

        # ── Authorized Devices -> IT Assets ─────────────────────
        if 'Authorized Devices' in wb.sheetnames:
            recs = _import_generic_table(wb['Authorized Devices'], ['#'], {
                'name': ['device name', 'model', 'name'],
                'type': ['type'],
                'serial': ['device serial #', 'device serial', 'serial #', 'serial'],
                'tracking_id': ['internal tracking #', 'internal tracking', 'tracking #', 'tracking id'],
                'owner': ['assignedto', 'assigned to', 'owner'],
                'os': ['os / vendor', 'os/vendor', 'operating system', 'os'],
            }, required_any=['name', 'type'])
            new_items = [{
                'id': str(uuid.uuid4()), 'name': r['name'] or r['type'], 'type': r['type'],
                'serial': r['serial'], 'tracking_id': r['tracking_id'], 'category': 'CUI',
                'location': '', 'owner': r['owner'], 'os': r['os'], 'ip': '',
                'notes': r['__extras__'], 'source': 'imported',
            } for r in recs]
            results['assets'] = _append_new(data['assets'], new_items, ['name', 'serial'])

        # ── Authorized Users -> Authorized Users ────────────────
        if 'Authorized Users' in wb.sheetnames:
            recs = _import_generic_table(wb['Authorized Users'], ['#'], {
                'user_name': ['user name'],
                'business_role': ['business roles'],
                'responsibilities': ['responsibilities'],
                'personnel_category': ['personnel category'],
                'functions': ['functions'],
                'privileged_raw': ['privileged user'],
                'cui_categories': ['authorized access to cui categories'],
                'post_cui_public': ['authorized to post business information to public spaces?'],
                'risk_level': ['risk level'],
                'assigned_devices': ['assigned devices'],
                'locations_raw': ['locations'],
                'work_site': ['authorized work sites'],
                'bldg_rm': ['bldg / rm'],
            }, required_any=['user_name'])
            new_items = []
            for r in recs:
                priv = r['privileged_raw']
                loc_first = r['locations_raw'].split('\n')[0].strip() if r['locations_raw'] else ''
                new_items.append({
                    'id': str(uuid.uuid4()), 'name': r['user_name'],
                    'business_role': r['business_role'],
                    'email': r['user_name'] if '@' in r['user_name'] else '',
                    'personnel_category': r['personnel_category'],
                    'responsibilities': r['responsibilities'], 'functions': r['functions'],
                    'is_privileged': bool(priv), 'privileged_server': [], 'privileged_firewall': [],
                    'privileged_network_device': [], 'privileged_ms365': [],
                    'pim_privileges': [l.strip() for l in priv.split('\n') if l.strip()],
                    'cui_categories': r['cui_categories'], 'post_cui_public': r['post_cui_public'] or 'No',
                    'risk_level': r['risk_level'], 'assigned_devices': r['assigned_devices'],
                    'location': loc_first, 'work_site': r['work_site'], 'bldg_rm': r['bldg_rm'],
                })
            results['authorized_users'] = _append_new(data['authorized_users'], new_items, ['name'])

        # ── Authorized Server Apps -> Server Apps ───────────────
        #    Mobile Whitelist       -> Mobile Whitelist (same layout)
        app_field_map = {
            'app_name': ['application'],
            'host': ['device name'],
            'app_type': ['type'],
            'deployment_model': ['on-prem / cloud'],
            # 'FedRAMP' gets camelCase-split to 'Fed RAMP' by _hdr_map()'s
            # normalizer (lowercase-then-uppercase boundary), so match both.
            'fedramp_approved': ['fedramp authorized?', 'fed ramp authorized?'],
            'version': ['version / release'],
        }
        for sheet_name, section_key in [('Authorized Server Apps', 'server_apps'),
                                          ('Mobile Whitelist', 'mobile_whitelist')]:
            if sheet_name in wb.sheetnames:
                recs = _import_generic_table(wb[sheet_name], ['#'], app_field_map, required_any=['app_name'])
                new_items = [{
                    'id': str(uuid.uuid4()), 'name': r['app_name'], 'app_type': r['app_type'],
                    'version': r['version'], 'deployment_model': r['deployment_model'],
                    'fedramp_approved': r['fedramp_approved'], 'host': r['host'],
                    'vendor': '', 'approved_by': '', 'purpose': '', 'notes': r['__extras__'],
                } for r in recs]
                results[section_key] = _append_new(data[section_key], new_items, ['name', 'host'])

        # ── Authorized System Processes -> System Processes ─────
        if 'Authorized System Processes' in wb.sheetnames:
            recs = _import_generic_table(wb['Authorized System Processes'], ['user / account name'], {
                'service_account': ['user / account name'],
                'name': ['process / service / agent name'],
                'path': ['executable path'],
                'host': ['system(s) running on'],
                'purpose': ['purpose / function'],
            }, required_any=['name', 'service_account'], skip_row_number_col=False)
            new_items = [{
                'id': str(uuid.uuid4()), 'name': r['name'] or r['service_account'],
                'path': r['path'], 'host': r['host'], 'service_account': r['service_account'],
                'purpose': r['purpose'], 'notes': r['__extras__'],
            } for r in recs]
            results['system_processes'] = _append_new(data['system_processes'], new_items, ['name', 'host'])

        # ── Non & Essential PPS -> Ports & Protocols ────────────
        if 'Non & Essential PPS' in wb.sheetnames:
            recs = _import_generic_table(wb['Non & Essential PPS'], ['system name'], {
                'system_name': ['system name'],
                'os': ['operating system'],
                'ppsdesc': ['essential program, function, port, protocol, service'],
                'justification': ['justification', 'method of restriction'],
            }, required_any=['system_name', 'ppsdesc'], skip_row_number_col=False)
            new_items = [{
                'id': str(uuid.uuid4()), 'port': '', 'protocol': '', 'service': r['ppsdesc'],
                'direction': '', 'source': f"{r['system_name']} ({r['os']})" if r['os'] else r['system_name'],
                'destination': '', 'justification': r['justification'],
            } for r in recs]
            results['ports_protocols'] = _append_new(data['ports_protocols'], new_items, ['service', 'source'])

        # ── Physical Security -> Physical Security ──────────────
        if 'Physical Security' in wb.sheetnames:
            recs = _import_generic_table(wb['Physical Security'], ['#'], {
                'locations_raw': ['locations'],
                'work_site': ['authorized work sites'],
                'pacs': ['physical access control devices'],
                'manager': ['building manager / fso'],
            }, required_any=['locations_raw'])
            new_items = []
            for r in recs:
                loc_first = r['locations_raw'].split('\n')[0].strip() if r['locations_raw'] else ''
                location = f"{loc_first} - {r['work_site']}".strip(' -') if r['work_site'] else loc_first
                new_items.append({
                    'id': str(uuid.uuid4()), 'location': location,
                    'device_type': 'Physical Security Controls',
                    'model': r['pacs'], 'serial': '', 'category': 'SPA',
                    'logs_access': '', 'cui_area': '', 'manager': r['manager'],
                    'notes': r['__extras__'],
                })
            results['physical_security'] = _append_new(data['physical_security'], new_items, ['location', 'manager'])

        # ── Workstation App Blacklist -> App Blacklist ──────────
        if 'Workstation App Blacklist' in wb.sheetnames:
            tracker_cats, tech_impl, notice = _parse_blacklist_sheet(wb['Workstation App Blacklist'])
            updated = 0
            for tc in tracker_cats:
                match = _find_matching_blacklist_category(data['blacklist_categories'], tc['name'])
                if match:
                    match['apps'] = tc['apps']
                    if tc['justification']:
                        match['justification'] = tc['justification']
                else:
                    key = tc['name'].strip().lower()
                    new_cat = {
                        'id': 'bl-' + re.sub(r'[^a-z0-9]+', '-', key).strip('-')[:40],
                        'name': tc['name'], 'justification': tc['justification'], 'apps': tc['apps'],
                    }
                    data['blacklist_categories'].append(new_cat)
                updated += 1
            if tech_impl:
                data['blacklist_tech_impl'] = tech_impl
            if notice:
                data['blacklist_notice'] = notice
            results['blacklist_categories'] = updated

        # Reference-only sheets — not asset records, already built into
        # Asset Inventory -> Reference (see docstring above).
        skipped_sheets = [s for s in ('Non-US Access', 'Contractor Risk Review',
                                       'Risk Heat Map', 'Risk Matrix') if s in wb.sheetnames]

        save_assets(data)
        total = sum(v for v in results.values() if isinstance(v, int))
        return jsonify({"success": True, "results": results, "skipped_sheets": skipped_sheets,
                        "imported": results.get('assets', 0), "total": total})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─────────────────────────────────────────────────────────────
# Intune / Entra Smart Import
# ─────────────────────────────────────────────────────────────

def _cell(v):
    """Normalize a cell value to a clean string, treating None/NaN/blanks as ''."""
    if v is None:
        return ''
    s = str(v).strip()
    return '' if s.lower() in ('none', 'nan', 'n/a', '-', '') else s

def _hdr_map(header_row):
    """Return {normalised_col_name: col_index} for a row of header values.
    Normalises camelCase (e.g. 'userPrincipalName') to space-separated lower
    ('user principal name') so lookups work regardless of export format.
    """
    import re as _re
    def norm(s):
        s = _re.sub(r'([a-z])([A-Z])', r'\1 \2', str(s or ''))
        return s.lower().strip()
    return {norm(h): i for i, h in enumerate(header_row)}

def _load_workbook_any(raw_bytes, filename=''):
    """
    Load a workbook from raw bytes regardless of format.
    Supports:
      - .xlsx / .xlsm  → opened with openpyxl directly
      - .csv            → parsed with csv module, wrapped in an openpyxl Workbook
    Both Intune and Entra portals export CSV by default; XLSX is available
    via "Export → Download" in some views.
    """
    fname_lower = (filename or '').lower()
    is_csv = fname_lower.endswith('.csv')

    # Also auto-detect CSV by trying xlsx first and falling back
    if not is_csv:
        try:
            return openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
        except Exception:
            is_csv = True  # not a valid zip/xlsx → treat as CSV

    if is_csv:
        import csv as _csv
        # Try UTF-8-BOM first (Azure/Intune exports often have BOM), then latin-1
        for encoding in ('utf-8-sig', 'utf-8', 'latin-1', 'cp1252'):
            try:
                text = raw_bytes.decode(encoding)
                break
            except (UnicodeDecodeError, LookupError):
                text = None
        if not text:
            raise ValueError("Could not decode CSV file — unexpected encoding.")

        reader = _csv.reader(io.StringIO(text))
        rows = list(reader)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        for row in rows:
            ws.append(row)
        return wb

    return openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)


def _detect_file_type(wb):
    """
    Examine sheet names and header rows to classify the workbook.
    Handles both space-separated headers (xlsx portal exports) and
    camelCase headers (CSV portal exports).
    """
    sheet = wb.active
    sheet_names_lower = [s.lower() for s in wb.sheetnames]

    # Our own tracker
    if any(s in sheet_names_lower for s in ['authorized devices', 'asset inventory']):
        return 'artifacts_tracker'

    # Read first non-empty row as header
    rows = []
    for row in sheet.iter_rows(values_only=True):
        clean = [_cell(v) for v in row]
        if any(clean):
            rows.append(clean)
        if len(rows) >= 2:
            break

    if not rows:
        return 'unknown'

    # Normalise header: lowercase, strip spaces so "DeviceName" == "device name" == "devicename"
    import re as _re
    def norm(s):
        # Insert space before uppercase runs so camelCase → "camel case"
        s = _re.sub(r'([a-z])([A-Z])', r'\1 \2', str(s))
        return s.lower().strip()

    raw_header = rows[0]
    header = [norm(h) for h in raw_header]
    header_str = ' '.join(header)

    # ── Intune Devices ────────────────────────────────────────
    # Intune device exports contain compliance, managed by, or device-specific IDs
    if any(k in header_str for k in [
        'managed by', 'compliance', 'enrollment date', 'enrolled date time',
        'intune device id', 'azure ad device id', 'enrolled date',
        'serial number',  # combined with OS = device context
    ]):
        # Make sure it's not a user file that happens to have serial number
        if any(k in header_str for k in ['device name', 'os', 'operating system', 'managed by', 'compliance']):
            return 'intune_devices'

    # ── Entra Role Assignments ────────────────────────────────
    # exportRoleAssignments has "assignedRole" / "assigned role" / "role type"
    if any(k in header_str for k in ['assigned role', 'role type', 'role name', 'directory role']):
        return 'entra_groups'

    # ── Intune/Entra user-device mapping ─────────────────────
    if 'user principal name' in header_str and any(k in header_str for k in ['device name', 'device count', 'registered devices']):
        return 'intune_users'

    # ── Entra ID Users (All Users / ID_exportUsers) ───────────
    # Has UPN + at least one user attribute
    if 'user principal name' in header_str and any(k in header_str for k in [
        'display name', 'job title', 'department', 'usage location',
        'account enabled', 'user type', 'assigned licenses',
    ]):
        return 'entra_users'

    # ── Entra Group Members ───────────────────────────────────
    if any(k in header_str for k in ['member type', 'group name', 'member of']):
        return 'entra_groups'

    # ── Entra MFA ─────────────────────────────────────────────
    if any(k in header_str for k in ['mfa status', 'mfa state', 'authentication method', 'default mfa method', 'per-user mfa']):
        return 'entra_mfa'

    # ── Entra Sign-in logs ────────────────────────────────────
    if any(k in header_str for k in ['sign-in risk', 'correlation id', 'sign-in status', 'conditional access']):
        return 'entra_signin'

    # ── Generic fallback ─────────────────────────────────────
    # UPN present at all → treat as some kind of user list
    if 'user principal name' in header_str:
        return 'entra_users'

    return 'unknown'


def _parse_intune_devices(wb, existing_assets):
    """
    Parse Intune All Devices / Managed Devices export.
    Returns (list_of_new_asset_dicts, added, skipped_dupes, warnings)
    Deduplicates on Serial Number → then on Device Name.
    """
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return [], 0, 0, []

    header = rows[0]
    h = _hdr_map(header)

    # Build dedup sets from existing
    existing_serials = {a.get('serial', '').strip().lower() for a in existing_assets if a.get('serial', '').strip()}
    existing_names   = {a.get('name', '').strip().lower()   for a in existing_assets if a.get('name', '').strip()}

    new_assets = []
    added = skipped = 0
    warnings = []

    # Column name variants Intune uses across portal versions
    name_cols    = ['device name', 'devicename', 'computer name', 'computername', 'name']
    serial_cols  = ['serial number', 'serialnumber', 'serial']
    os_cols      = ['operating system', 'os', 'operatingsystem', 'os version']
    type_cols    = ['device type', 'devicetype', 'ownership', 'ownership type']
    owner_cols   = ['primary user upn', 'primary user email', 'primary user', 'userupn', 'user upn']
    ip_cols      = ['ip address', 'ip address (ipv4)', 'ipaddress', 'last sync ip']
    mgmt_cols    = ['managed by', 'management type', 'managementtype']
    enroll_cols  = ['enrollment date', 'enrolled date', 'enrollmentdate']
    compliance_cols = ['compliance', 'compliance status', 'compliancestatus']
    model_cols   = ['model', 'device model', 'manufacturer']

    def find_col(candidates):
        for c in candidates:
            if c in h:
                return h[c]
        return None

    i_name       = find_col(name_cols)
    i_serial     = find_col(serial_cols)
    i_os         = find_col(os_cols)
    i_type       = find_col(type_cols)
    i_owner      = find_col(owner_cols)
    i_ip         = find_col(ip_cols)
    i_mgmt       = find_col(mgmt_cols)
    i_enroll     = find_col(enroll_cols)
    i_compliance = find_col(compliance_cols)
    i_model      = find_col(model_cols)

    if i_name is None:
        warnings.append("Could not find a 'Device Name' column. Check that this is an Intune Devices export.")

    for row in rows[1:]:
        if not row or not any(_cell(v) for v in row):
            continue

        get = lambda i: _cell(row[i]) if i is not None and i < len(row) else ''

        name    = get(i_name)
        serial  = get(i_serial)
        os_val  = get(i_os)
        owner   = get(i_owner)
        ip      = get(i_ip)
        model   = get(i_model)
        mgmt    = get(i_mgmt)
        compliance = get(i_compliance)
        enroll  = get(i_enroll)
        raw_type = get(i_type)

        if not name and not serial:
            continue

        # Dedup
        if serial and serial.lower() in existing_serials:
            skipped += 1
            continue
        if name and name.lower() in existing_names:
            skipped += 1
            continue

        # Infer device type
        device_type = 'Workstation (Desktop)'
        os_lower = os_val.lower()
        name_lower = name.lower()
        raw_lower = raw_type.lower()
        if 'server' in os_lower or 'server' in name_lower:
            device_type = 'Server (Physical)'
        elif 'ios' in os_lower or 'iphone' in os_lower or 'ipad' in os_lower:
            device_type = 'Mobile Device (Smartphone)'
        elif 'android' in os_lower:
            device_type = 'Mobile Device (Smartphone)'
        elif 'mac' in os_lower or 'macos' in os_lower or 'osx' in os_lower:
            device_type = 'Laptop'
        elif 'windows' in os_lower:
            device_type = 'Laptop'

        notes_parts = []
        if mgmt:      notes_parts.append(f"Managed by: {mgmt}")
        if compliance: notes_parts.append(f"Compliance: {compliance}")
        if enroll:    notes_parts.append(f"Enrolled: {enroll}")
        if model:     notes_parts.append(f"Model: {model}")

        asset = {
            'id': str(uuid.uuid4()),
            'name': name,
            'type': device_type,
            'category': 'CUI',
            'location': '',
            'owner': owner,
            'ip': ip,
            'os': os_val,
            'serial': serial,
            'tracking_id': '',
            'encrypted': '',
            'notes': '; '.join(notes_parts),
            'source': 'intune',
        }
        new_assets.append(asset)
        if serial: existing_serials.add(serial.lower())
        if name:   existing_names.add(name.lower())
        added += 1

    return new_assets, added, skipped, warnings


def _parse_entra_users(wb, existing_users):
    """
    Parse Entra ID (Azure AD) Users export AND generic Intune user list.
    Returns (list_of_new_user_dicts, added, updated, skipped, warnings)
    Deduplicates on UPN (email), then merges data when UPN already exists.
    """
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return [], 0, 0, 0, []

    header = rows[0]
    h = _hdr_map(header)

    # Build UPN → user dict reference map (covers both existing and newly-staged users)
    upn_to_user = {}
    for u in existing_users:
        email = (u.get('email') or '').strip().lower()
        if email:
            upn_to_user[email] = u

    # Column name variants across Entra export versions (camelCase normalised by _hdr_map)
    upn_cols      = ['user principal name', 'upn', 'email', 'email address', 'mail', 'sign-in name']
    name_cols     = ['display name', 'name', 'full name']
    given_cols    = ['first name', 'given name', 'firstname']
    surname_cols  = ['last name', 'surname', 'lastname', 'family name']
    title_cols    = ['job title', 'title', 'role']
    dept_cols     = ['department', 'dept']
    type_cols     = ['user type', 'account type']
    status_cols   = ['account enabled', 'account status', 'enabled', 'status', 'sign-in status']
    group_cols    = ['assigned role', 'assigned roles', 'directory roles', 'directory role',
                     'role name', 'group', 'groups', 'member of', 'role type']
    phone_cols    = ['phone number', 'mobile', 'phone']
    location_cols = ['usage location', 'country', 'city', 'office location', 'office']
    mfa_cols      = ['mfa status', 'mfa state', 'per-user mfa status', 'strong authentication state']
    license_cols  = ['assigned licenses', 'licenses', 'license']

    def find_col(candidates):
        for c in candidates:
            if c in h:
                return h[c]
        return None

    i_upn      = find_col(upn_cols)
    i_name     = find_col(name_cols)
    i_given    = find_col(given_cols)
    i_surname  = find_col(surname_cols)
    i_title    = find_col(title_cols)
    i_dept     = find_col(dept_cols)
    i_type     = find_col(type_cols)
    i_status   = find_col(status_cols)
    i_group    = find_col(group_cols)
    i_phone    = find_col(phone_cols)
    i_location = find_col(location_cols)
    i_mfa      = find_col(mfa_cols)
    i_license  = find_col(license_cols)

    warnings = []
    if i_upn is None:
        warnings.append("Could not find a 'User Principal Name' / email column. Check that this is an Entra ID or Intune Users export.")

    new_users = []
    added = updated = skipped = 0

    for row in rows[1:]:
        if not row or not any(_cell(v) for v in row):
            continue

        get = lambda i: _cell(row[i]) if i is not None and i < len(row) else ''

        upn    = get(i_upn)
        name   = get(i_name)
        given  = get(i_given)
        sname  = get(i_surname)
        title  = get(i_title)
        dept   = get(i_dept)
        utype  = get(i_type)
        status = get(i_status)
        group  = get(i_group)
        phone  = get(i_phone)
        loc    = get(i_location)
        mfa    = get(i_mfa)
        lic    = get(i_license)

        if not upn:
            continue

        # Build display name
        if not name and (given or sname):
            name = f"{given} {sname}".strip()

        upn_lower = upn.lower()

        # Merge into already-seen user (existing or newly staged in this call)
        if upn_lower in upn_to_user:
            u = upn_to_user[upn_lower]
            changed = False
            if name and not u.get('name'):
                u['name'] = name; changed = True
            if title and not u.get('business_role'):
                u['business_role'] = title; changed = True
            if phone and not u.get('phone'):
                u['phone'] = phone; changed = True
            if loc and not u.get('location'):
                u['location'] = loc; changed = True
            if dept and not u.get('responsibilities'):
                u['responsibilities'] = dept; changed = True
            # Accumulate roles/groups — append, don't overwrite
            if group:
                existing_funcs = u.get('functions', '')
                existing_groups = set(g.strip() for g in existing_funcs.split('|') if g.strip()) if existing_funcs else set()
                if group not in existing_groups:
                    existing_groups.add(group)
                    u['functions'] = ' | '.join(sorted(existing_groups))
                    changed = True
            if mfa or lic or status:
                notes_parts = []
                if mfa:    notes_parts.append(f"MFA: {mfa}")
                if lic:    notes_parts.append(f"Licenses: {lic}")
                if status: notes_parts.append(f"Status: {status}")
                extra = '; '.join(notes_parts)
                if extra and extra not in (u.get('notes') or ''):
                    u['notes'] = ((u.get('notes') or '') + '; ' + extra).strip('; ')
                    changed = True
            # Only count as "updated" if this was a pre-existing user (not one we just added)
            if changed and u in existing_users:
                updated += 1
            else:
                skipped += 1
            continue

        # Determine personnel category
        personnel_cat = 'Full-Time Employee (FTE)'
        utype_lower = utype.lower()
        if 'guest' in utype_lower or 'external' in utype_lower or 'b2b' in utype_lower:
            personnel_cat = 'Third-Party Vendor'

        notes_parts = []
        if mfa:    notes_parts.append(f"MFA: {mfa}")
        if lic:    notes_parts.append(f"Licenses: {lic}")
        if status: notes_parts.append(f"Status: {status}")

        user = {
            'id': str(uuid.uuid4()),
            'name': name or upn,
            'business_role': title,
            'email': upn,
            'personnel_category': personnel_cat,
            'responsibilities': dept,
            'functions': group or '',
            'is_privileged': False,
            'privileged_server': [],
            'privileged_firewall': [],
            'privileged_network_device': [],
            'privileged_ms365': [],
            'pim_privileges': [],
            'cui_categories': '',
            'post_cui_public': 'No',
            'risk_level': '',
            'assigned_devices': '',
            'location': loc,
            'work_site': '',
            'bldg_rm': '',
            'notes': '; '.join(notes_parts),
            'source': 'entra',
        }
        new_users.append(user)
        upn_to_user[upn_lower] = user  # register for same-call dedup
        added += 1

    return new_users, added, updated, skipped, warnings


@app.route('/api/import-intune-entra', methods=['POST'])
def api_import_intune_entra():
    """
    Accept one or more Excel (xlsx) or CSV files from Intune / Entra exports.
    Auto-detects file type and merges into the asset / user inventory.
    Returns a detailed result for each file.
    """
    files = request.files.getlist('files')
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    data = load_assets()
    results = []
    total_devices_added = 0
    total_users_added = 0
    total_users_updated = 0
    all_warnings = []

    for f in files:
        fname = f.filename
        try:
            raw = f.read()
            wb = _load_workbook_any(raw, fname)
            ftype = _detect_file_type(wb)

            if ftype == 'intune_devices':
                new_assets, added, skipped, warnings = _parse_intune_devices(wb, data['assets'])
                data['assets'].extend(new_assets)
                save_assets(data)
                total_devices_added += added
                results.append({
                    "file": fname, "type": "Intune Devices",
                    "added": added, "skipped_dupes": skipped,
                    "warnings": warnings,
                })
                all_warnings.extend(warnings)

            elif ftype in ('entra_users', 'intune_users', 'entra_groups', 'entra_mfa', 'entra_signin'):
                new_users, added, updated, skipped, warnings = _parse_entra_users(wb, data['authorized_users'])
                data['authorized_users'].extend(new_users)
                save_assets(data)
                total_users_added += added
                total_users_updated += updated
                type_label = {
                    'entra_users': 'Entra ID Users',
                    'intune_users': 'Intune User List',
                    'entra_groups': 'Entra Group Members',
                    'entra_mfa': 'Entra MFA Report',
                    'entra_signin': 'Entra Sign-in Log',
                }.get(ftype, 'Entra Users')
                results.append({
                    "file": fname, "type": type_label,
                    "added": added, "updated": updated, "skipped_dupes": skipped,
                    "warnings": warnings,
                })
                all_warnings.extend(warnings)

            elif ftype == 'artifacts_tracker':
                # Fall back to original import logic for our own tracker format
                imported = 0
                if 'Authorized Devices' in wb.sheetnames:
                    ws = wb['Authorized Devices']
                    existing_names = {a.get('name','').lower() for a in data['assets']}
                    header_row = None
                    for row in ws.iter_rows(values_only=True):
                        if row and _cell(row[0]) == '#':
                            header_row = row; continue
                        if header_row and row and _cell(row[0]) and _cell(row[0]) not in ['#', 'Ex.']:
                            try:
                                int(_cell(row[0]))
                                name = _cell(row[1])
                                if name and name.lower() not in existing_names:
                                    data['assets'].append({
                                        'id': str(uuid.uuid4()),
                                        'name': name,
                                        'type': _cell(row[2]),
                                        'serial': _cell(row[3]),
                                        'tracking_id': _cell(row[4]),
                                        'category': 'CUI', 'location': '',
                                        'owner': '', 'os': '', 'ip': '',
                                        'notes': '', 'encrypted': '',
                                        'source': 'artifacts_tracker',
                                    })
                                    existing_names.add(name.lower())
                                    imported += 1
                            except (ValueError, TypeError):
                                pass
                save_assets(data)
                results.append({"file": fname, "type": "CMMC Artifacts Tracker", "added": imported, "skipped_dupes": 0, "warnings": []})
                total_devices_added += imported

            else:
                results.append({
                    "file": fname, "type": "Unknown",
                    "added": 0, "skipped_dupes": 0,
                    "warnings": [f"Could not identify the file format. Supported: Intune All Devices, Entra ID Users, Entra Group Members, Entra MFA report, or CMMC Artifacts Tracker."],
                })

        except Exception as e:
            results.append({"file": fname, "type": "Error", "added": 0, "error": str(e), "warnings": []})

    return jsonify({
        "success": True,
        "results": results,
        "total_devices_added": total_devices_added,
        "total_users_added": total_users_added,
        "total_users_updated": total_users_updated,
    })

# ─────────────────────────────────────────────────────────────
# Export: Asset Tracker XLSX
# ─────────────────────────────────────────────────────────────
@app.route('/export/xlsx')
def export_xlsx():
    data = load_assets()
    ssp = load_ssp()
    wb = build_asset_xlsx(data, ssp)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    org = ssp.get('org_abbr', 'Customer') or 'Customer'
    return send_file(buf, download_name=f"{org}_CMMC_L2_Asset_Tracker.xlsx",
                     as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ─────────────────────────────────────────────────────────────
# Export: SSP DOCX
# ─────────────────────────────────────────────────────────────
@app.route('/export/ssp')
def export_ssp():
    ssp = load_ssp()
    assets = load_assets()
    doc = build_ssp_docx(ssp, assets)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    org = ssp.get('org_abbr', 'Customer') or 'Customer'
    return send_file(buf, download_name=f"{org}_SSP_v{ssp.get('version','1.0')}.docx",
                     as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

# ─────────────────────────────────────────────────────────────
# Static image serving
# ─────────────────────────────────────────────────────────────
_UPLOAD_ALLOWED_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp'}
_UPLOAD_ALLOWED_PREFIXES = ('logo_', 'network_diagram_', 'physical_diagram_', 'cui_flow_diagram_')

@app.route('/uploads/<filename>')
def serve_upload(filename):
    # Only ever serve the fixed set of images this app itself writes to
    # UPLOADS_DIR — filenames are always "<prefix>_<8 hex chars><ext>" (see
    # images()) — never trust the URL segment as a path.
    stem = Path(filename).stem.lower()
    ext = Path(filename).suffix.lower()
    prefix = next((p for p in _UPLOAD_ALLOWED_PREFIXES if stem.startswith(p)), None)
    valid_suffix = prefix and re.fullmatch(r'[0-9a-f]{8}', stem[len(prefix):])
    if not valid_suffix or ext not in _UPLOAD_ALLOWED_EXTS:
        return jsonify({"error": "Not found"}), 404
    fpath = (UPLOADS_DIR / filename).resolve()
    if UPLOADS_DIR.resolve() not in fpath.parents or not fpath.is_file():
        return jsonify({"error": "Not found"}), 404
    return send_file(fpath)

# ─────────────────────────────────────────────────────────────
# Excel Builder
# ─────────────────────────────────────────────────────────────
def set_border(cell, style='thin'):
    bd = Side(style=style, color='CCCCCC')
    cell.border = Border(left=bd, right=bd, top=bd, bottom=bd)

def header_cell(cell, text, color='1F3864', font_color='FFFFFF', bold=True, size=10):
    cell.value = text
    cell.font = Font(bold=bold, color=font_color, size=size, name='Arial')
    cell.fill = PatternFill('solid', fgColor=color)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    set_border(cell)

def data_cell(cell, text='', bold=False, color=None, wrap=True, align='left'):
    cell.value = text
    cell.font = Font(bold=bold, size=9, name='Arial')
    cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    if color:
        cell.fill = PatternFill('solid', fgColor=color)
    set_border(cell)

def build_asset_xlsx(data, ssp):
    wb = openpyxl.Workbook()
    org = ssp.get('org_name', '[Organization]')
    org_abbr = ssp.get('org_abbr', 'Customer')
    
    # ── Sheet 1: Asset Inventory ──────────────────────────────
    ws = wb.active
    ws.title = "Asset Inventory"
    ws.sheet_view.showGridLines = False
    
    # Title row
    ws.merge_cells('A1:P1')
    ws['A1'].value = f"CMMC L2 Asset Inventory – {org}"
    ws['A1'].font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    ws['A1'].fill = PatternFill('solid', fgColor='1F3864')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # Sub-header: category legend
    ws.merge_cells('A2:P2')
    ws['A2'].value = ("Categories: CUI=CUI Asset | SPA=Security Protection Asset | "
                      "CRMA=Contractor Risk Managed | SA_GFE/IoT/OT/RI/TE=Specialized | OOS=Out-of-Scope")
    ws['A2'].font = Font(italic=True, size=8, color='444444', name='Arial')
    ws['A2'].fill = PatternFill('solid', fgColor='D9E1F2')
    ws['A2'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[2].height = 20

    headers = ['#', 'Asset Name / Hostname', 'Device Type', 'CMMC Category', 'Category Description',
               'Serial Number', 'Internal Tracking ID', 'IP Address', 'OS / Firmware',
               'Location / Site', 'Owner / Custodian', 'CUI Processed?', 'Encrypted (DAR)?',
               'In Scope?', 'Notes / Justification', 'Last Updated']
    col_widths = [4, 28, 20, 12, 38, 18, 18, 14, 18, 16, 16, 12, 12, 10, 32, 14]
    
    ws.row_dimensions[3].height = 36
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        header_cell(ws.cell(3, c), h)
        ws.column_dimensions[get_column_letter(c)].width = w

    assets = data.get('assets', [])
    alt_colors = ['FFFFFF', 'EEF2FF']
    for r, asset in enumerate(assets, 4):
        cat = asset.get('category', 'CUI')
        in_scope = 'No' if cat == 'OOS' else 'Yes'
        cui_proc = 'Yes' if cat == 'CUI' else ('Possible' if cat == 'CRMA' else 'No')
        row_color = alt_colors[r % 2]
        vals = [
            r - 3,
            asset.get('name', ''),
            asset.get('type', ''),
            cat,
            ASSET_CATEGORIES.get(cat, ''),
            asset.get('serial', ''),
            asset.get('tracking_id', ''),
            asset.get('ip', ''),
            asset.get('os', ''),
            asset.get('location', ''),
            asset.get('owner', ''),
            cui_proc,
            asset.get('encrypted', ''),
            in_scope,
            asset.get('notes', ''),
            asset.get('modified', asset.get('created', ''))[:10] if asset.get('modified') or asset.get('created') else '',
        ]
        ws.row_dimensions[r].height = 18
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c)
            bold = c == 1
            # Color-code category
            if c == 4:
                cat_colors = {'CUI': 'FFD966', 'SPA': 'A9D18E', 'CRMA': 'F4B942',
                              'OOS': 'BFBFBF', 'SA_GFE': 'BDD7EE', 'SA_IoT': 'BDD7EE',
                              'SA_OT': 'BDD7EE', 'SA_RI': 'BDD7EE', 'SA_TE': 'BDD7EE'}
                data_cell(cell, v, bold=True, color=cat_colors.get(cat, 'FFFFFF'), align='center')
            elif c == 14:
                data_cell(cell, v, color='C6EFCE' if v == 'Yes' else 'FFCCCC', align='center')
            elif c == 12:
                data_cell(cell, v, color='FFD966' if v == 'Possible' else ('C6EFCE' if v == 'No' else 'FFD7B3'), align='center')
            else:
                data_cell(cell, v, bold=bold, color=row_color)

    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}3"

    # ── Sheet 2: Physical Security ────────────────────────────
    ws2 = wb.create_sheet("Physical Security")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells('A1:K1')
    ws2['A1'].value = f"Physical Security Asset Inventory – {org}"
    ws2['A1'].font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    ws2['A1'].fill = PatternFill('solid', fgColor='1F3864')
    ws2['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[1].height = 30

    ph_headers = ['#', 'Site / Location', 'Device Type', 'Manufacturer / Model',
                  'Serial / Asset Tag', 'CMMC Category', 'Manages/Logs Access?',
                  'CUI Area Protected?', 'Building Manager / FSO', 'Integration Notes', 'Last Updated']
    ph_widths = [4, 18, 24, 22, 18, 12, 16, 16, 20, 30, 14]
    ws2.row_dimensions[2].height = 36
    for c, (h, w) in enumerate(zip(ph_headers, ph_widths), 1):
        header_cell(ws2.cell(2, c), h)
        ws2.column_dimensions[get_column_letter(c)].width = w

    physical = data.get('physical_security', [])
    for r, item in enumerate(physical, 3):
        ws2.row_dimensions[r].height = 18
        row_color = alt_colors[r % 2]
        vals = [
            r - 2,
            item.get('location', ''),
            item.get('device_type', ''),
            item.get('model', ''),
            item.get('serial', ''),
            item.get('category', 'SPA'),
            item.get('logs_access', ''),
            item.get('cui_area', ''),
            item.get('manager', ''),
            item.get('notes', ''),
            '',
        ]
        for c, v in enumerate(vals, 1):
            cell = ws2.cell(r, c)
            if c == 6:
                data_cell(cell, v, color='A9D18E', align='center')
            else:
                data_cell(cell, v, bold=(c == 1), color=row_color)

    ws2.freeze_panes = 'A3'

    # ── Sheet 3: Authorized Users ─────────────────────────────
    ws_users = wb.create_sheet("Authorized Users")
    ws_users.sheet_view.showGridLines = False
    ws_users.merge_cells('A1:R1')
    ws_users['A1'].value = f"Authorized Users – {org}"
    ws_users['A1'].font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    ws_users['A1'].fill = PatternFill('solid', fgColor='1F3864')
    ws_users['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws_users.row_dimensions[1].height = 30

    user_headers = ['#', 'User Name', 'Business Role', 'User Email', 'Personnel Category',
                     'Responsibilities', 'Functions / Role', 'Privileged User?',
                     'Privileged: Server', 'Privileged: Firewall', 'Privileged: Network Device',
                     'Privileged: MS365 Roles', 'PIM Privileges',
                     'Authorized CUI Categories', 'Post CUI to Public?', 'Risk Level',
                     'Assigned Devices', 'Location', 'Authorized Work Site', 'Bldg/Rm (ACL)']
    user_widths = [4, 20, 22, 26, 20, 32, 28, 12, 22, 22, 22, 26, 28, 20, 16, 14, 22, 22, 18, 16]
    ws_users.row_dimensions[2].height = 36
    for c, (h, w) in enumerate(zip(user_headers, user_widths), 1):
        header_cell(ws_users.cell(2, c), h)
        ws_users.column_dimensions[get_column_letter(c)].width = w

    for r, u in enumerate(data.get('authorized_users', []), 3):
        row_color = alt_colors[r % 2]
        is_priv = u.get('is_privileged', False)
        vals = [r-2, u.get('name',''), u.get('business_role',''), u.get('email',''),
                u.get('personnel_category',''), u.get('responsibilities',''), u.get('functions',''),
                'Yes' if is_priv else 'No',
                ', '.join(u.get('privileged_server', []) or []),
                ', '.join(u.get('privileged_firewall', []) or []),
                ', '.join(u.get('privileged_network_device', []) or []),
                ', '.join(u.get('privileged_ms365', []) or []),
                ', '.join(u.get('pim_privileges', []) or []),
                u.get('cui_categories',''), u.get('post_cui_public',''),
                u.get('risk_level',''), u.get('assigned_devices',''), u.get('location',''),
                u.get('work_site',''), u.get('bldg_rm','')]
        ws_users.row_dimensions[r].height = 18
        for c, v in enumerate(vals, 1):
            cell = ws_users.cell(r, c)
            if c == 8:  # Privileged User?
                data_cell(cell, v, color='FFD966' if v == 'Yes' else 'F2F2F2', align='center')
            elif c == 15:  # Post CUI to Public?
                data_cell(cell, v, color='FFCCCC' if v == 'Yes' else 'C6EFCE', align='center')
            else:
                data_cell(cell, v, bold=(c == 1), color=row_color)
    ws_users.freeze_panes = 'A3'
    if data.get('authorized_users'):
        ws_users.auto_filter.ref = f"A2:{get_column_letter(len(user_headers))}2"

    # ── Sheet 4: Authorized Server Apps ───────────────────────
    ws_sa = wb.create_sheet("Authorized Server Apps")
    ws_sa.sheet_view.showGridLines = False
    ws_sa.merge_cells('A1:K1')
    ws_sa['A1'].value = f"Authorized Server Applications – {org}"
    ws_sa['A1'].font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    ws_sa['A1'].fill = PatternFill('solid', fgColor='1F3864')
    ws_sa['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws_sa.row_dimensions[1].height = 30

    sa_headers = ['#', 'Application Name', 'Application Type', 'Version', 'On-Prem / Cloud',
                  'FedRAMP Approved', 'Server / Host', 'Vendor / Publisher',
                  'Business Purpose', 'Approved By', 'Notes']
    sa_widths = [4, 26, 22, 16, 14, 16, 20, 24, 38, 20, 30]
    ws_sa.row_dimensions[2].height = 30
    for c, (h, w) in enumerate(zip(sa_headers, sa_widths), 1):
        header_cell(ws_sa.cell(2, c), h)
        ws_sa.column_dimensions[get_column_letter(c)].width = w
    for r, a in enumerate(data.get('server_apps', []), 3):
        row_color = alt_colors[r % 2]
        vals = [r-2, a.get('name',''), a.get('app_type',''), a.get('version',''),
                a.get('deployment_model',''), a.get('fedramp_approved',''), a.get('host',''),
                a.get('vendor',''), a.get('purpose',''), a.get('approved_by',''), a.get('notes','')]
        ws_sa.row_dimensions[r].height = 18
        for c, v in enumerate(vals, 1):
            cell = ws_sa.cell(r, c)
            if c == 6:  # FedRAMP Approved
                data_cell(cell, v, color='C6EFCE' if v == 'Yes' else ('FFCCCC' if v == 'No' else 'F2F2F2'), align='center')
            else:
                data_cell(cell, v, bold=(c == 1), color=row_color)
    ws_sa.freeze_panes = 'A3'

    # ── Sheet 4b: Mobile Whitelist ──────────────────────────────
    ws_mw = wb.create_sheet("Mobile Whitelist")
    ws_mw.sheet_view.showGridLines = False
    ws_mw.merge_cells('A1:K1')
    ws_mw['A1'].value = f"Mobile Application Whitelist – {org}"
    ws_mw['A1'].font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    ws_mw['A1'].fill = PatternFill('solid', fgColor='1F3864')
    ws_mw['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws_mw.row_dimensions[1].height = 30

    mw_headers = ['#', 'Application Name', 'Application Type', 'Version', 'On-Prem / Cloud',
                  'FedRAMP Approved', 'Device / Host', 'Vendor / Publisher',
                  'Business Purpose', 'Approved By', 'Notes']
    mw_widths = [4, 26, 22, 16, 14, 16, 20, 24, 38, 20, 30]
    ws_mw.row_dimensions[2].height = 30
    for c, (h, w) in enumerate(zip(mw_headers, mw_widths), 1):
        header_cell(ws_mw.cell(2, c), h)
        ws_mw.column_dimensions[get_column_letter(c)].width = w
    for r, a in enumerate(data.get('mobile_whitelist', []), 3):
        row_color = alt_colors[r % 2]
        vals = [r-2, a.get('name',''), a.get('app_type',''), a.get('version',''),
                a.get('deployment_model',''), a.get('fedramp_approved',''), a.get('host',''),
                a.get('vendor',''), a.get('purpose',''), a.get('approved_by',''), a.get('notes','')]
        ws_mw.row_dimensions[r].height = 18
        for c, v in enumerate(vals, 1):
            cell = ws_mw.cell(r, c)
            if c == 6:  # FedRAMP Approved
                data_cell(cell, v, color='C6EFCE' if v == 'Yes' else ('FFCCCC' if v == 'No' else 'F2F2F2'), align='center')
            else:
                data_cell(cell, v, bold=(c == 1), color=row_color)
    ws_mw.freeze_panes = 'A3'

    # ── Sheet 5: Authorized System Processes ──────────────────
    ws_proc = wb.create_sheet("Authorized System Processes")
    ws_proc.sheet_view.showGridLines = False
    ws_proc.merge_cells('A1:G1')
    ws_proc['A1'].value = f"Authorized System Processes – {org}"
    ws_proc['A1'].font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    ws_proc['A1'].fill = PatternFill('solid', fgColor='1F3864')
    ws_proc['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws_proc.row_dimensions[1].height = 30

    proc_headers = ['#', 'Process Name', 'Executable Path', 'Host / Server',
                     'Purpose', 'Started By / Service Account', 'Notes']
    proc_widths = [4, 22, 42, 20, 32, 28, 28]
    ws_proc.row_dimensions[2].height = 30
    for c, (h, w) in enumerate(zip(proc_headers, proc_widths), 1):
        header_cell(ws_proc.cell(2, c), h)
        ws_proc.column_dimensions[get_column_letter(c)].width = w
    for r, p in enumerate(data.get('system_processes', []), 3):
        row_color = alt_colors[r % 2]
        vals = [r-2, p.get('name',''), p.get('path',''), p.get('host',''),
                p.get('purpose',''), p.get('service_account',''), p.get('notes','')]
        ws_proc.row_dimensions[r].height = 18
        for c, v in enumerate(vals, 1):
            data_cell(ws_proc.cell(r, c), v, bold=(c == 1), color=row_color)
    ws_proc.freeze_panes = 'A3'

    # ── Sheet 6: Authorized Ports & Protocols ──────────────────
    ws_ports = wb.create_sheet("Authorized Ports & Protocols")
    ws_ports.sheet_view.showGridLines = False
    ws_ports.merge_cells('A1:H1')
    ws_ports['A1'].value = f"Authorized Ports & Protocols – {org}"
    ws_ports['A1'].font = Font(bold=True, size=14, color='FFFFFF', name='Arial')
    ws_ports['A1'].fill = PatternFill('solid', fgColor='1F3864')
    ws_ports['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws_ports.row_dimensions[1].height = 30

    port_headers = ['#', 'Port', 'Protocol', 'Service / Application', 'Direction',
                     'Source', 'Destination', 'Business Justification']
    port_widths = [4, 10, 12, 24, 14, 22, 22, 40]
    ws_ports.row_dimensions[2].height = 30
    for c, (h, w) in enumerate(zip(port_headers, port_widths), 1):
        header_cell(ws_ports.cell(2, c), h)
        ws_ports.column_dimensions[get_column_letter(c)].width = w
    for r, p in enumerate(data.get('ports_protocols', []), 3):
        row_color = alt_colors[r % 2]
        vals = [r-2, p.get('port',''), p.get('protocol',''), p.get('service',''),
                p.get('direction',''), p.get('source',''), p.get('destination',''), p.get('justification','')]
        ws_ports.row_dimensions[r].height = 18
        for c, v in enumerate(vals, 1):
            data_cell(ws_ports.cell(r, c), v, bold=(c == 1), color=row_color)
    ws_ports.freeze_panes = 'A3'

    # ── Sheet 7: Category Reference ───────────────────────────
    ws3 = wb.create_sheet("Category Reference")
    ws3.sheet_view.showGridLines = False
    ws3.merge_cells('A1:C1')
    ws3['A1'].value = "CMMC L2 Asset Category Reference (Per CMMC Scoping Guide v2.13)"
    ws3['A1'].font = Font(bold=True, size=12, color='FFFFFF', name='Arial')
    ws3['A1'].fill = PatternFill('solid', fgColor='1F3864')
    ws3['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws3.row_dimensions[1].height = 28

    for c, h in enumerate(['Code', 'Category Name', 'Description'], 1):
        header_cell(ws3.cell(2, c), h)
    ws3.column_dimensions['A'].width = 12
    ws3.column_dimensions['B'].width = 34
    ws3.column_dimensions['C'].width = 55

    cat_details = [
        ('CUI', 'CUI Asset', 'Processes, stores, or transmits CUI. Assessed against ALL CMMC L2 requirements.'),
        ('SPA', 'Security Protection Asset', 'Provides security functions/capabilities (firewalls, SIEM, VPN). Assessed against relevant CMMC L2 requirements.'),
        ('CRMA', 'Contractor Risk Managed Asset', 'Can process CUI but controlled by policy/procedure. In-scope; limited assessment if well-documented in SSP.'),
        ('SA_GFE', 'Specialized – GFE', 'Government Furnished Equipment. Documented in SSP; managed via risk-based practices.'),
        ('SA_IoT', 'Specialized – IoT/IIoT', 'Internet of Things / Industrial IoT devices. Documented in SSP; managed via risk-based practices.'),
        ('SA_OT', 'Specialized – OT', 'Operational Technology (industrial control, building management, PACS). Documented in SSP.'),
        ('SA_RI', 'Specialized – Restricted IS', 'Restricted information systems. Documented in SSP; managed via risk-based practices.'),
        ('SA_TE', 'Specialized – Test Equipment', 'Test equipment. Documented in SSP; managed via risk-based practices.'),
        ('OOS', 'Out-of-Scope Asset', 'Cannot process/store/transmit CUI AND does not provide security protections. Physically/logically separated.'),
    ]
    row_colors_cat = {'CUI': 'FFF2CC', 'SPA': 'E2EFDA', 'CRMA': 'FCE4D6',
                      'SA_GFE': 'DDEEFF', 'SA_IoT': 'DDEEFF', 'SA_OT': 'DDEEFF',
                      'SA_RI': 'DDEEFF', 'SA_TE': 'DDEEFF', 'OOS': 'F2F2F2'}
    for r, (code, name, desc) in enumerate(cat_details, 3):
        rc = row_colors_cat.get(code, 'FFFFFF')
        for c, v in enumerate([code, name, desc], 1):
            data_cell(ws3.cell(r, c), v, bold=(c <= 2), color=rc)
        ws3.row_dimensions[r].height = 24

    return wb

# ─────────────────────────────────────────────────────────────
# SSP DOCX Builder
# ─────────────────────────────────────────────────────────────
def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    return p

def add_para(doc, text='', bold=False, italic=False, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p

def fmt_body_para(p):
    """Apply the CMMC Controls Status narrative paragraph formatting: left
    aligned, no indentation, 6pt spacing before/after, single line spacing —
    matching the reference document's own body-text paragraph style."""
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.left_indent = Inches(0)
    pf.right_indent = Inches(0)
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    return p

def add_bullets(doc, text):
    """Split a newline-separated textarea value into 'List Bullet' paragraphs,
    skipping blank lines."""
    for line in (text or '').split('\n'):
        line = line.strip()
        if line:
            doc.add_paragraph(line, style='List Bullet')

def add_table_row(table, cells, bold_first=False, shade=None):
    row = table.add_row()
    for i, (cell, text) in enumerate(zip(row.cells, cells)):
        cell.text = str(text) if text is not None else ''
        run = cell.paragraphs[0].runs
        if run:
            run[0].bold = (bold_first and i == 0)
            run[0].font.size = Pt(9)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        if shade:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), shade)
            tcPr.append(shd)
    return row

def shade_row(row, fill_color):
    for cell in row.cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill_color)
        tcPr.append(shd)

def set_col_widths(table, widths_inches):
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_inches):
                cell.width = Inches(widths_inches[i])

def sanitize_text(text):
    """Strip characters that are illegal in XML/DOCX (control chars, null bytes, etc.)."""
    if not text:
        return ''
    # Replace common PDF encoding artifacts with their correct chars
    replacements = {
        '\x02': '-',   # soft hyphen / encoding artifact
        '\x00': '',    # null byte
        '\x01': '',
        '\x03': '',
        '\x04': '',
        '\x05': '',
        '\x06': '',
        '\x07': '',
        '\x08': '',
        '\x0b': ' ',   # vertical tab → space
        '\x0c': '',    # form feed
        '\x0e': '',
        '\x0f': '',
        '\x10': '',
        '\x11': '',
        '\x12': '',
        '\x13': '',
        '\x14': '',
        '\x15': '',
        '\x16': '',
        '\x17': '',
        '\x18': '',
        '\x19': '',
        '\x1a': '',
        '\x1b': '',
        '\x1c': '',
        '\x1d': '',
        '\x1e': '',
        '\x1f': '',
        '\ufffd': '',  # replacement character
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)
    # Final pass: remove any remaining chars with codepoints in illegal XML range
    return ''.join(c for c in text if ord(c) >= 0x20 or c in ('\t', '\n', '\r'))

def embed_diagram_list(doc, filenames, placeholder_text, width_inches=5.5):
    """
    Embed every image in a multi-diagram slot (network/physical/CUI-flow),
    centered, one after another — falls back to an italic placeholder
    paragraph when nothing has been uploaded for that slot.
    """
    existing = [f for f in (filenames or []) if f and (UPLOADS_DIR / f).exists()]
    if not existing:
        p = doc.add_paragraph(placeholder_text)
        p.runs[0].italic = True
        p.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        return
    for i, fname in enumerate(existing):
        doc.add_picture(str(UPLOADS_DIR / fname), width=Inches(width_inches))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if i < len(existing) - 1:
            doc.add_paragraph()

def build_ssp_docx(ssp, assets_data):
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1.25)

    org = ssp.get('org_name') or '[OSC]'
    org_abbr = ssp.get('org_abbr') or '[OSC]'
    mssp = ssp.get('mssp_name') or '[MSSP]'
    ver = ssp.get('version', '1.0')
    updated = ssp.get('last_updated', datetime.now().strftime('%m.%d.%Y'))

    # ── Cover Page ────────────────────────────────────────────
    # Logo
    logo_path = UPLOADS_DIR / ssp['images'].get('logo', '') if ssp['images'].get('logo') else None
    if logo_path and logo_path.exists():
        doc.add_picture(str(logo_path), width=Inches(2.5))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p = doc.add_paragraph('[Organization Logo]')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].italic = True
        p.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    doc.add_paragraph()
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(f"{org_abbr}, LLC. - System Security Plan (SSP)")
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

    doc.add_paragraph()
    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.add_run(f"Prepared By: {ssp.get('prepared_by', '')}")
    
    ver_p = doc.add_paragraph()
    ver_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ver_p.add_run(f"VERSION {ver}  |  {updated}")

    doc.add_page_break()

    # ── 1. System Identification ──────────────────────────────
    add_heading(doc, '1. System Identification', 1)

    add_heading(doc, 'Information Impact Categorization', 2)
    tbl = doc.add_table(rows=1, cols=5)
    tbl.style = 'Table Grid'
    hdr = tbl.rows[0]
    for cell, h in zip(hdr.cells, ['Category', 'Confidentiality', 'Integrity', 'Availability', 'Authority']):
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        shade_row(hdr, '1F3864')
        for r in hdr.cells:
            for run in r.paragraphs[0].runs:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for row_data in [(r.get('category',''), r.get('confidentiality',''), r.get('integrity',''),
                      r.get('availability',''), r.get('authority',''))
                     for r in ssp.get('impact_categorization', [])]:
        add_table_row(tbl, row_data)

    doc.add_paragraph()

    # Responsible Organization
    add_heading(doc, 'Responsible Organization', 2)
    rp = doc.add_paragraph()
    rp.add_run(f"Organization Name: ").bold = True
    rp.add_run(org)
    
    # System info table
    tbl2 = doc.add_table(rows=4, cols=2)
    tbl2.style = 'Table Grid'
    info_rows = [
        ('System Name', ssp.get('system_name', org)),
        ('System Type(s)', ssp.get('system_type', '')),
        ('System Categorization', ssp.get('system_categorization', 'CUI, FCI')),
        ('System Unique Identifier', ssp.get('system_uid', 'IT Infrastructure')),
    ]
    for row, (label, val) in zip(tbl2.rows, info_rows):
        row.cells[0].text = label
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[1].text = val

    doc.add_paragraph()
    
    # Personnel
    for section_title, person_key in [('Information Owner', 'info_owner'),
                                        ('System Owner', 'system_owner'),
                                        ('System Security Officer', 'sso')]:
        add_heading(doc, section_title, 2)
        person = ssp.get(person_key, {})
        tbl_p = doc.add_table(rows=4, cols=2)
        tbl_p.style = 'Table Grid'
        for row, (lbl, key) in zip(tbl_p.rows, [('Name', 'name'),('Title', 'title'),('Email', 'email'),('Phone', 'phone')]):
            row.cells[0].text = lbl
            row.cells[0].paragraphs[0].runs[0].bold = True
            row.cells[1].text = person.get(key, '')
        doc.add_paragraph()

    # User counts
    add_heading(doc, 'Number of End Users and Privileged Users', 2)
    eu_p = doc.add_paragraph()
    eu_p.add_run('End Users: ').bold = True
    eu_p.add_run(ssp.get('end_users', ''))
    priv_p = doc.add_paragraph()
    priv_p.add_run('Privileged Users: ').bold = True
    priv_p.add_run(ssp.get('privileged_users', ''))

    # CUI Categories
    add_heading(doc, 'General Description of CUI Categories', 2)
    doc.add_paragraph(ssp.get('cui_categories', ''))

    # NARA CUI Categories
    nara_categories = ssp.get('nara_cui_categories', [])
    if nara_categories:
        add_heading(doc, 'NARA CUI Categories', 2)
        nara_intro = ssp.get('nara_cui_intro', '') or ''
        doc.add_paragraph(nara_intro.replace('[OSC]', org_abbr) + ':')
        for cat in nara_categories:
            doc.add_paragraph(cat, style='List Bullet')

    locations = ssp.get('locations', [])

    # CUI Data Flow Environment
    add_heading(doc, 'CUI Data Flow Environment', 2)
    doc.add_paragraph(ssp.get('cui_data_flow', ''))
    embed_diagram_list(doc, ssp['images'].get('cui_flow_diagram'), '[CUI Data Flow Diagram – Insert diagram here]')
    doc.add_paragraph()

    # System Environment
    add_heading(doc, 'System Environment', 2)
    doc.add_paragraph(ssp.get('system_environment', ''))
    for i, loc in enumerate(locations):
        add_heading(doc, loc, 3)
        placeholder = '[High-level Detail Network Diagram]' if i == 0 else '[Detailed Network Diagram]'
        p = doc.add_paragraph(placeholder)
        p.runs[0].italic = True
        p.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    embed_diagram_list(doc, ssp['images'].get('network_diagram'), '[Network Diagram – Insert diagram here]')
    doc.add_paragraph()
    # Preserve pre-redesign content so nothing already written is lost on upgrade
    if ssp.get('network_desc', '').strip():
        doc.add_paragraph(ssp['network_desc'])

    # Physical Security
    add_heading(doc, 'Physical Security', 2)
    phys_by_loc = ssp.get('physical_security_by_location', {})
    if locations:
        for loc in locations:
            add_heading(doc, loc, 3)
            doc.add_paragraph(phys_by_loc.get(loc, ''))
    elif ssp.get('physical_security_desc', '').strip():
        doc.add_paragraph(ssp['physical_security_desc'])
    embed_diagram_list(doc, ssp['images'].get('physical_diagram'), '[Floor Diagram highlighting all PADS and PACS]')
    doc.add_paragraph()

    # Physical Security Assets from inventory
    phys_assets = assets_data.get('physical_security', [])
    if phys_assets:
        add_heading(doc, 'Physical Security Technology', 3)
        pa_tbl = doc.add_table(rows=1, cols=5)
        pa_tbl.style = 'Table Grid'
        pa_hdr = pa_tbl.rows[0]
        shade_row(pa_hdr, '1F3864')
        for cell, h in zip(pa_hdr.cells, ['Site', 'Device Type', 'Model', 'Category', 'Notes']):
            cell.text = h
            if cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].bold = True
                cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        for pa in phys_assets:
            add_table_row(pa_tbl, [pa.get('location',''), pa.get('device_type',''),
                                   pa.get('model',''), pa.get('category','SPA'), pa.get('notes','')])
        doc.add_paragraph()

    # ── CMMC Level 2 Asset Categories ─────────────────────────
    add_heading(doc, 'CMMC Level 2 Asset Categories', 2)

    add_heading(doc, 'CUI Assets', 3)
    add_para(doc, 'Workstations', bold=True, size=11)
    doc.add_paragraph(ssp.get('workstation_desc', ''))

    add_para(doc, 'Servers (or Out of Scope)', bold=True, size=11)
    doc.add_paragraph(ssp.get('server_desc', ''))
    for st in ssp.get('server_types', []):
        doc.add_paragraph(st, style='List Bullet')

    server_assets = [a for a in assets_data.get('assets', []) if 'server' in a.get('type', '').lower() or 'domain controller' in a.get('type', '').lower()]
    if server_assets:
        sv_tbl = doc.add_table(rows=1, cols=5)
        sv_tbl.style = 'Table Grid'
        sv_hdr = sv_tbl.rows[0]
        shade_row(sv_hdr, '2F5496')
        for cell, h in zip(sv_hdr.cells, ['Asset Name', 'Type', 'OS', 'IP Address', 'Category']):
            cell.text = h
            if cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].bold = True
                cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        for a in server_assets:
            add_table_row(sv_tbl, [a.get('name',''), a.get('type',''), a.get('os',''), a.get('ip',''), a.get('category','')])
    doc.add_paragraph()

    add_heading(doc, 'Contractor Risk Managed Assets (CRMA)', 3)
    add_para(doc, 'Mobile Devices', bold=True, size=11)
    doc.add_paragraph(ssp.get('crma_mobile_desc', ''))

    add_heading(doc, 'Specialized Assets (SA)', 3)
    if ssp.get('specialized_assets_desc', '').strip():
        doc.add_paragraph(ssp['specialized_assets_desc'])
    else:
        na = doc.add_paragraph('None identified.')
        na.runs[0].italic = True
        na.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    add_heading(doc, 'Security Protection Assets (SPA)', 3)
    spa_by_loc = ssp.get('spa_network_by_location', {})
    if locations:
        add_para(doc, 'Network / Firewall', bold=True, size=11)
        for i, loc in enumerate(locations):
            label = f"{loc} (Main Site)" if i == 0 else loc
            add_para(doc, label, bold=True, size=10)
            doc.add_paragraph(spa_by_loc.get(loc, ''))
    add_para(doc, 'Wireless Access Points (or Out of Scope)', bold=True, size=11)
    doc.add_paragraph(ssp.get('spa_wireless_desc', ''))
    add_para(doc, 'Whitelisted USB Storage Devices', bold=True, size=11)
    doc.add_paragraph(ssp.get('spa_usb_desc', ''))
    add_para(doc, 'Microsoft 365 GCC[-High]', bold=True, size=11)
    doc.add_paragraph(ssp.get('cloud_apps_desc', ''))

    add_heading(doc, 'Out-of-Scope Assets (OOS)', 3)
    if ssp.get('out_of_scope_assets_desc', '').strip():
        doc.add_paragraph(ssp['out_of_scope_assets_desc'])
    else:
        na2 = doc.add_paragraph('None identified.')
        na2.runs[0].italic = True
        na2.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    # ── CUI System Inventory ──────────────────────────────────
    add_heading(doc, 'CUI System Inventory', 2)
    inv_p = doc.add_paragraph()
    inv_p.add_run(f"{org_abbr} asset inventory is maintained in the attached CMMC L2 Asset Tracker spreadsheet. "
                  "The following table provides a summary of in-scope assets by category.")

    all_assets = assets_data.get('assets', [])
    if all_assets:
        inv_tbl = doc.add_table(rows=1, cols=6)
        inv_tbl.style = 'Table Grid'
        inv_hdr = inv_tbl.rows[0]
        shade_row(inv_hdr, '1F3864')
        for cell, h in zip(inv_hdr.cells, ['#', 'Asset Name', 'Type', 'Category', 'Location', 'In Scope']):
            cell.text = h
            if cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].bold = True
                cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        for i, a in enumerate(all_assets, 1):
            in_scope = 'No' if a.get('category') == 'OOS' else 'Yes'
            add_table_row(inv_tbl, [i, a.get('name',''), a.get('type',''), a.get('category',''), a.get('location',''), in_scope])
        doc.add_paragraph()

    doc.add_page_break()

    # ── Controls (Administrative Controls / Technical Controls narrative) ──
    add_heading(doc, 'Controls', 2)
    doc.add_paragraph(ssp.get('controls_intro', ''))

    add_heading(doc, 'Administrative Controls', 3)
    doc.add_paragraph(ssp.get('admin_controls_intro', ''))

    add_para(doc, 'Policies, Standards, and Procedures', bold=True, size=11)
    doc.add_paragraph(ssp.get('policies_intro', ''))
    add_para(doc, 'Policies', bold=True, size=10)
    add_bullets(doc, ssp.get('policies_list', ''))
    add_para(doc, 'Standards', bold=True, size=10)
    add_bullets(doc, ssp.get('standards_list', ''))
    add_para(doc, 'Procedures and Playbooks', bold=True, size=10)
    add_bullets(doc, ssp.get('procedures_list', ''))

    add_para(doc, 'Access Control', bold=True, size=11)
    doc.add_paragraph(ssp.get('access_control_desc', ''))

    add_para(doc, 'Asset Management', bold=True, size=11)
    doc.add_paragraph(ssp.get('asset_mgmt_desc', ''))

    add_para(doc, 'Audit and Assessment', bold=True, size=11)
    doc.add_paragraph(ssp.get('audit_assessment_desc', ''))

    add_para(doc, 'Authentication', bold=True, size=11)
    doc.add_paragraph(ssp.get('authentication_desc', ''))

    add_para(doc, 'Compliance Management', bold=True, size=11)
    doc.add_paragraph(ssp.get('compliance_mgmt_desc', ''))

    add_para(doc, 'Incident Response Program', bold=True, size=11)
    doc.add_paragraph(ssp.get('incident_response_desc', ''))

    add_para(doc, 'Personnel Security', bold=True, size=11)
    doc.add_paragraph(ssp.get('personnel_security_desc', ''))

    add_para(doc, 'Risk & Security Analysis', bold=True, size=11)
    doc.add_paragraph(ssp.get('risk_analysis_admin_desc', ''))

    add_para(doc, 'Roles and Responsibilities', bold=True, size=11)
    doc.add_paragraph(ssp.get('roles_resp_admin_intro', ''))
    add_bullets(doc, ssp.get('roles_resp_srm_items', ''))
    doc.add_paragraph(ssp.get('roles_resp_admin_closing', ''))

    add_para(doc, 'Security Awareness Training', bold=True, size=11)
    doc.add_paragraph(ssp.get('security_awareness_admin_intro', ''))
    add_bullets(doc, ssp.get('security_awareness_admin_items', ''))

    add_para(doc, 'Vulnerability Management', bold=True, size=11)
    doc.add_paragraph(ssp.get('vulnerability_mgmt_desc', ''))

    add_heading(doc, 'Technical Controls', 3)
    doc.add_paragraph(ssp.get('tech_controls_intro', ''))

    add_para(doc, 'Audit and Logging', bold=True, size=11)
    doc.add_paragraph(ssp.get('audit_logging_desc', ''))

    add_para(doc, 'Configuration Management', bold=True, size=11)
    doc.add_paragraph(ssp.get('config_mgmt_intro', ''))
    add_bullets(doc, ssp.get('config_mgmt_items', ''))

    add_para(doc, 'Encryption', bold=True, size=11)
    doc.add_paragraph(ssp.get('encryption_intro', ''))
    add_bullets(doc, ssp.get('encryption_items', ''))
    doc.add_paragraph(ssp.get('encryption_closing', ''))

    add_para(doc, 'Malware and System Intrusion Protection', bold=True, size=11)
    doc.add_paragraph(ssp.get('malware_protection_desc', ''))

    add_para(doc, 'Media Handling', bold=True, size=11)
    doc.add_paragraph(ssp.get('media_handling_intro', ''))
    add_bullets(doc, ssp.get('media_handling_items', ''))

    add_para(doc, 'Network Security', bold=True, size=11)
    doc.add_paragraph(ssp.get('network_security_desc', ''))
    doc.add_paragraph(ssp.get('network_security_wireless_desc', ''))

    add_para(doc, 'Physical Security', bold=True, size=11)
    doc.add_paragraph(ssp.get('physical_security_tech_desc', ''))

    add_para(doc, 'Risk & Security Analysis', bold=True, size=11)
    doc.add_paragraph(ssp.get('risk_analysis_tech_desc', ''))

    add_para(doc, 'Roles and Responsibilities', bold=True, size=11)
    doc.add_paragraph(ssp.get('roles_resp_tech_desc', ''))

    add_para(doc, 'Security Awareness Training', bold=True, size=11)
    doc.add_paragraph(ssp.get('security_awareness_tech_desc', ''))

    add_para(doc, 'System Update', bold=True, size=11)
    doc.add_paragraph(ssp.get('system_update_desc', ''))

    # CMMC Controls Status section (below) gets its own narrower left/right
    # margins — 0.8" vs. the document's 1.25" default — via a dedicated
    # section break, matching the reference document's page setup. The
    # section reverts to the default margins again right before Version
    # History (see the matching add_section() call further down).
    ctrl_section = doc.add_section(WD_SECTION.NEW_PAGE)
    ctrl_section.top_margin = Inches(1)
    ctrl_section.bottom_margin = Inches(1)
    ctrl_section.left_margin = Inches(0.8)
    ctrl_section.right_margin = Inches(0.8)

    # ── 2. Controls Summary ───────────────────────────────────
    add_heading(doc, '2. Controls', 1)
    controls_status = ssp.get('controls_status', {})
    
    add_heading(doc, 'CMMC L2 Controls Implementation Status', 2)
    ctrl_tbl = doc.add_table(rows=1, cols=4)
    ctrl_tbl.style = 'Table Grid'
    c_hdr = ctrl_tbl.rows[0]
    shade_row(c_hdr, '1F3864')
    for cell, h in zip(c_hdr.cells, ['Control ID', 'Control Name', 'Status', 'Notes']):
        cell.text = h
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    status_colors = {
        'Implemented': 'C6EFCE',
        'Planned': 'FFEB9C',
        'Not Applicable': 'BDD7EE',
        'Not Implemented': 'FFC7CE',
    }
    
    for ctrl in CMMC_CONTROLS:
        ctrl_id = ctrl['id']
        status = controls_status.get(ctrl_id, 'Implemented')
        row = add_table_row(ctrl_tbl, [ctrl_id, ctrl['name'], status, ''])
        # Color status cell
        for cell in row.cells[2:3]:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), status_colors.get(status, 'FFFFFF'))
            tcPr.append(shd)

    doc.add_page_break()

    # ── 3. CMMC Compliance ────────────────────────────────────
    add_heading(doc, '3. CMMC Compliance', 1)
    note = doc.add_paragraph()
    note.add_run('Note: ').bold = True
    note.add_run('The requirements below are sourced from the Cybersecurity Maturity Model Certification 2.0 – '
                 'Level 2 Framework and derived from NIST Special Publication 800-171 Rev. 2. '
                 'Assessment objectives are drawn from the CMMC Assessment Guide v2.13 (NIST SP 800-171A).').italic = True
    fmt_body_para(note)

    controls_narratives = ssp.get('controls_narratives', {})

    # Group controls by domain
    domains = {}
    for ctrl in CMMC_CONTROLS:
        d = ctrl['domain']
        domains.setdefault(d, []).append(ctrl)

    for domain_name, domain_ctrls in domains.items():
        add_heading(doc, domain_name, 2)
        for ctrl in domain_ctrls:
            ctrl_id = ctrl['id']
            status = controls_status.get(ctrl_id, 'Implemented')

            # Control ID + Name heading
            p = doc.add_paragraph()
            run = p.add_run(f"{ctrl_id} \u2013 {ctrl['name'].upper()}")
            run.bold = True
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

            # Status checkboxes
            status_p = doc.add_paragraph()
            statuses = ['Implemented', 'Planned to be Implemented', 'Not Applicable']
            status_p.add_run('  '.join(
                f"[{'✓' if status == s else ' '}] {s}" for s in statuses
            ))
            status_p.runs[0].font.size = Pt(9)

            # Requirement description
            desc_p = doc.add_paragraph(ctrl.get('description', ''))
            desc_p.runs[0].font.size = Pt(10) if desc_p.runs else None
            fmt_body_para(desc_p)

            # Assessment objectives + narratives
            objectives = ctrl.get('objectives', [])
            ctrl_narratives = controls_narratives.get(ctrl_id, {})

            if objectives:
                obj_label = doc.add_paragraph()
                obj_run = obj_label.add_run('Assessment Objectives (Enforcement Criteria):')
                obj_run.bold = True
                obj_run.font.size = Pt(9)
                obj_run.font.color.rgb = RGBColor(0x2F, 0x54, 0x96)
                fmt_body_para(obj_label)

                for letter, obj_text in objectives:
                    # Sanitize: strip control chars that cause XML errors
                    obj_text_clean = sanitize_text(obj_text)
                    # Objective letter + text
                    obj_p = doc.add_paragraph(style='List Bullet')
                    lb = obj_p.add_run(f'[{letter}]  ')
                    lb.bold = True
                    lb.font.size = Pt(9)
                    lb.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
                    ob = obj_p.add_run(f'Determine if: {obj_text_clean}.')
                    ob.italic = True
                    ob.font.size = Pt(9)

                    # Narrative for this objective — saved text wins; otherwise fall
                    # back to the same default boilerplate shown on the Controls page.
                    narrative_raw = ctrl_narratives.get(letter, '').strip() or \
                                     CONTROLS_NARRATIVES.get(ctrl_id, {}).get(letter, '').strip()
                    narrative = sanitize_text(narrative_raw)
                    narr_p = doc.add_paragraph()
                    fmt_body_para(narr_p)
                    narr_run = narr_p.add_run(f'Enforcement Narrative [{letter}]: ')
                    narr_run.bold = True
                    narr_run.font.size = Pt(9)
                    narr_run.font.color.rgb = RGBColor(0x37, 0x56, 0x23)
                    if narrative:
                        narr_p.add_run(narrative).font.size = Pt(9)
                    else:
                        blank_run = narr_p.add_run('[No narrative entered]')
                        blank_run.font.size = Pt(9)
                        blank_run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
                        blank_run.italic = True

            doc.add_paragraph()  # spacing

    # ── 4. Version History ────────────────────────────────────
    # Revert to the document's default margins now that the CMMC Controls
    # Status section (with its narrower 0.8" left/right margins) is done.
    vh_section = doc.add_section(WD_SECTION.NEW_PAGE)
    vh_section.top_margin = Inches(1)
    vh_section.bottom_margin = Inches(1)
    vh_section.left_margin = Inches(1.25)
    vh_section.right_margin = Inches(1.25)
    add_heading(doc, '4. Document Version History', 1)
    vh_tbl = doc.add_table(rows=1, cols=4)
    vh_tbl.style = 'Table Grid'
    vh_hdr = vh_tbl.rows[0]
    shade_row(vh_hdr, '1F3864')
    for cell, h in zip(vh_hdr.cells, ['Version', 'Date', 'Author', 'Description']):
        cell.text = h
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for vh in ssp.get('version_history', []):
        add_table_row(vh_tbl, [vh.get('version',''), vh.get('date',''), vh.get('author',''), vh.get('description','')])

    return doc

# ─────────────────────────────────────────────────────────────
# CMMC L2 Controls Reference
# ─────────────────────────────────────────────────────────────
CMMC_CONTROLS = [
    # Access Control
    {"id":'AC.L2-3.1.1',"name":'Authorized Access Control',"domain":'3.1 Access Control (AC)',"description":'Limit system access to authorized users, processes acting on behalf of authorized users, and devices (including other systems).',"objectives":[['a', 'authorized users are identified'], ['b', 'processes acting on behalf of authorized users are identified'], ['c', 'devices (and other systems) authorized to connect to the system are identified'], ['d', 'system access is limited to authorized users'], ['e', 'system access is limited to processes acting on behalf of authorized users; and'], ['f', 'system access is limited to authorized devices (including other systems)']]},
    {"id":'AC.L2-3.1.2',"name":'Transaction & Function Control',"domain":'3.1 Access Control (AC)',"description":'Limit system access to the types of transactions and functions that authorized users are permitted to execute.',"objectives":[['a', 'the types of transactions and functions that authorized users are permitted to execute are defined; and'], ['b', 'system access is limited to the defined types of transactions and functions for authorized users']]},
    {"id":'AC.L2-3.1.3',"name":'Control CUI Flow',"domain":'3.1 Access Control (AC)',"description":'Control the flow of CUI in accordance with approved authorizations.',"objectives":[['a', 'information flow control policies are defined'], ['b', 'methods and enforcement mechanisms for controlling the flow of CUI are defined'], ['c', 'designated sources and destinations (e.g., networks, individuals, and devices) for CUI within the system and between interconnected systems are identified'], ['d', 'authorizations for controlling the flow of CUI are defined; and'], ['e', 'approved authorizations for controlling the flow of CUI are enforced']]},
    {"id":'AC.L2-3.1.4',"name":'Separation of Duties',"domain":'3.1 Access Control (AC)',"description":'Separate the duties of individuals to reduce the risk of malevolent activity without collusion.',"objectives":[['a', 'the duties of individuals requiring separation are defined'], ['b', 'responsibilities for duties that require separation are assigned to separate individuals; and'], ['c', 'access privileges that enable individuals to exercise the duties that require separation are granted to separate individuals']]},
    {"id":'AC.L2-3.1.5',"name":'Least Privilege',"domain":'3.1 Access Control (AC)',"description":'Employ the principle of least privilege, including for specific security functions and privileged accounts.',"objectives":[['a', 'privileged accounts are identified'], ['b', 'access to privileged accounts is authorized in accordance with the principle of least privilege'], ['c', 'security functions are identified; and'], ['d', 'access to security functions is authorized in accordance with the principle of least privilege']]},
    {"id":'AC.L2-3.1.6',"name":'Non-Privileged Account Use',"domain":'3.1 Access Control (AC)',"description":'Use non-privileged accounts or roles when accessing nonsecurity functions.',"objectives":[['a', 'nonsecurity functions are identified; and'], ['b', 'users are required to use non-privileged accounts or roles when accessing nonsecurity functions']]},
    {"id":'AC.L2-3.1.7',"name":'Privileged Functions',"domain":'3.1 Access Control (AC)',"description":'Prevent non-privileged users from executing privileged functions and capture the execution of such functions in audit logs.',"objectives":[['a', 'privileged functions are defined'], ['b', 'non-privileged users are defined'], ['c', 'non-privileged users are prevented from executing privileged functions; and'], ['d', 'the execution of privileged functions is captured in audit logs']]},
    {"id":'AC.L2-3.1.8',"name":'Unsuccessful Logon Attempts',"domain":'3.1 Access Control (AC)',"description":'Limit unsuccessful logon attempts.',"objectives":[['a', 'the means of limiting unsuccessful logon attempts is defined; and'], ['b', 'the defined means of limiting unsuccessful logon attempts is implemented']]},
    {"id":'AC.L2-3.1.9',"name":'Privacy & Security Notices',"domain":'3.1 Access Control (AC)',"description":'Provide privacy and security notices consistent with applicable CUI rules.',"objectives":[['a', 'privacy and security notices required by CUI-specified rules are identified, consistent, and associated with the specific CUI category; and'], ['b', 'privacy and security notices are displayed']]},
    {"id":'AC.L2-3.1.10',"name":'Session Lock',"domain":'3.1 Access Control (AC)',"description":'Use session lock with pattern-hiding displays to prevent access and viewing of data after a period of inactivity.',"objectives":[['a', 'the period of inactivity after which the system initiates a session lock is defined'], ['b', 'access to the system and viewing of data is prevented by initiating a session lock after the defined period of inactivity; and'], ['c', 'previously visible information is concealed via a pattern-hiding display after the defined period of inactivity']]},
    {"id":'AC.L2-3.1.11',"name":'Session Termination',"domain":'3.1 Access Control (AC)',"description":'Terminate (automatically) a user session after a defined condition.',"objectives":[['a', 'conditions requiring a user session to terminate are defined; and'], ['b', 'a user session is automatically terminated after any of the defined conditions occur']]},
    {"id":'AC.L2-3.1.12',"name":'Control Remote Access',"domain":'3.1 Access Control (AC)',"description":'Monitor and control remote access sessions.',"objectives":[['a', 'remote access sessions are permitted'], ['b', 'the types of permitted remote access are identified'], ['c', 'remote access sessions are controlled; and'], ['d', 'remote access sessions are monitored']]},
    {"id":'AC.L2-3.1.13',"name":'Remote Access Confidentiality',"domain":'3.1 Access Control (AC)',"description":'Employ cryptographic mechanisms to protect the confidentiality of remote access sessions.',"objectives":[['a', 'cryptographic mechanisms to protect the confidentiality of remote access sessions are identified; and'], ['b', 'cryptographic mechanisms to protect the confidentiality of remote access sessions are implemented']]},
    {"id":'AC.L2-3.1.14',"name":'Remote Access Routing',"domain":'3.1 Access Control (AC)',"description":'Route remote access via managed access control points.',"objectives":[['a', 'managed access control points are identified and implemented; and'], ['b', 'remote access is routed through managed network access control points']]},
    {"id":'AC.L2-3.1.15',"name":'Privileged Remote Access',"domain":'3.1 Access Control (AC)',"description":'Authorize remote execution of privileged commands and remote access to security-relevant information.',"objectives":[['a', 'privileged commands authorized for remote execution are identified'], ['b', 'security-relevant information authorized to be accessed remotely is identified'], ['c', 'the execution of the identified privileged commands via remote access is authorized; and'], ['d', 'access to the identified security-relevant information via remote access is authorized']]},
    {"id":'AC.L2-3.1.16',"name":'Wireless Access Authorization',"domain":'3.1 Access Control (AC)',"description":'Authorize wireless access prior to allowing such connections.',"objectives":[['a', 'wireless access points are identified; and'], ['b', 'wireless access is authorized prior to allowing such connections']]},
    {"id":'AC.L2-3.1.17',"name":'Wireless Access Protection',"domain":'3.1 Access Control (AC)',"description":'Protect wireless access using authentication and encryption.',"objectives":[['a', 'wireless access to the system is protected using authentication; and'], ['b', 'wireless access to the system is protected using encryption']]},
    {"id":'AC.L2-3.1.18',"name":'Mobile Device Connection',"domain":'3.1 Access Control (AC)',"description":'Control connection of mobile devices.',"objectives":[['a', 'mobile devices that process, store, or transmit CUI are identified'], ['b', 'mobile device connections are authorized; and'], ['c', 'mobile device connections are monitored and logged']]},
    {"id":'AC.L2-3.1.19',"name":'Encrypt CUI on Mobile',"domain":'3.1 Access Control (AC)',"description":'Encrypt CUI on mobile devices and mobile computing platforms.',"objectives":[['a', 'mobile devices and mobile computing platforms that process, store, or transmit CUI are identified; and'], ['b', 'encryption is employed to protect CUI on identified mobile devices and mobile computing platforms']]},
    {"id":'AC.L2-3.1.20',"name":'External Connections',"domain":'3.1 Access Control (AC)',"description":'Verify and control/limit connections to and use of external systems.',"objectives":[['a', 'connections to external systems are identified'], ['b', 'the use of external systems is identified'], ['c', 'connections to external systems are verified'], ['d', 'the use of external systems is verified'], ['e', 'connections to external systems are controlled/limited; and'], ['f', 'the use of external systems is controlled/limited']]},
    {"id":'AC.L2-3.1.21',"name":'Portable Storage Use',"domain":'3.1 Access Control (AC)',"description":'Limit use of portable storage devices on external systems.',"objectives":[['a', 'the use of portable storage devices containing CUI on external systems is identified and documented'], ['b', 'limits on the use of portable storage devices containing CUI on external systems are defined; and'], ['c', 'the use of portable storage devices containing CUI on external systems is limited as defined']]},
    {"id":'AC.L2-3.1.22',"name":'Control Public Information',"domain":'3.1 Access Control (AC)',"description":'Control CUI posted or processed on publicly accessible systems.',"objectives":[['a', 'individuals authorized to post or process information on publicly accessible systems are identified'], ['b', 'procedures to ensure CUI is not posted or processed on publicly accessible systems are identified'], ['c', 'a review process is in place prior to posting of any content to publicly accessible systems'], ['d', 'content on publicly accessible systems is reviewed to ensure that it does not include CUI; and'], ['e', 'mechanisms are in place to remove and address improper posting of CUI']]},
    # Awareness & Training
    {"id":'AT.L2-3.2.1',"name":'Role-Based Risk Awareness',"domain":'3.2 Awareness and Training (AT)',"description":'Ensure that managers, systems administrators, and users are made aware of the security risks associated with their activities and applicable policies, standards, and procedures.',"objectives":[['a', 'security risks associated with organizational activities involving CUI are identified'], ['b', 'policies, standards, and procedures related to the security of the system are identified'], ['c', 'managers, systems administrators, and users of the system are made aware of the security risks associated with their activities; and'], ['d', 'managers, systems administrators, and users of the system are made aware of the applicable policies, standards, and procedures related to the security of the system']]},
    {"id":'AT.L2-3.2.2',"name":'Role-Based Training',"domain":'3.2 Awareness and Training (AT)',"description":'Ensure that personnel are trained to carry out their assigned information security-related duties and responsibilities.',"objectives":[['a', 'information security-related duties, roles, and responsibilities are defined'], ['b', 'information security-related duties, roles, and responsibilities are assigned to designated personnel; and'], ['c', 'personnel are adequately trained to carry out their assigned information security-related duties, roles, and responsibilities']]},
    {"id":'AT.L2-3.2.3',"name":'Insider Threat Awareness',"domain":'3.2 Awareness and Training (AT)',"description":'Provide security awareness training on recognizing and reporting potential indicators of insider threat.',"objectives":[['a', 'potential indicators associated with insider threats are identified; and'], ['b', 'security awareness training on recognizing and reporting potential indicators of insider threat is provided to managers and employees']]},
    # Audit & Accountability
    {"id":'AU.L2-3.3.1',"name":'System Auditing',"domain":'3.3 Audit and Accountability (AU)',"description":'Create and retain system audit logs and records to enable monitoring, analysis, investigation, and reporting of unlawful or unauthorized system activity.',"objectives":[['a', 'audit logs needed (i.e., event types to be logged) to enable the monitoring, analysis, investigation, and reporting of unlawful or unauthorized system activity are specified'], ['b', 'the content of audit records needed to support monitoring, analysis, investigation, and reporting of unlawful or unauthorized system activity is defined'], ['c', 'audit records are created (generated)'], ['d', 'audit records, once created, contain the defined content'], ['e', 'retention requirements for audit records are defined; and'], ['f', 'audit records are retained as defined']]},
    {"id":'AU.L2-3.3.2',"name":'User Accountability',"domain":'3.3 Audit and Accountability (AU)',"description":'Ensure that the actions of individual system users can be uniquely traced to those users so they can be held accountable.',"objectives":[['a', 'the content of the audit records needed to support the ability to uniquely trace users to their actions is defined; and'], ['b', 'audit records, once created, contain the defined content']]},
    {"id":'AU.L2-3.3.3',"name":'Event Review',"domain":'3.3 Audit and Accountability (AU)',"description":'Review and update logged events.',"objectives":[['a', 'a process for determining when to review logged events is defined'], ['b', 'event types being logged are reviewed in accordance with the defined review process; and'], ['c', 'event types being logged are updated based on the review']]},
    {"id":'AU.L2-3.3.4',"name":'Audit Failure Alerting',"domain":'3.3 Audit and Accountability (AU)',"description":'Alert in the event of an audit logging process failure.',"objectives":[['a', 'personnel or roles to be alerted in the event of an audit logging process failure are identified'], ['b', 'types of audit logging process failures for which alert will be generated are defined; and'], ['c', 'identified personnel or roles are alerted in the event of an audit logging process failure']]},
    {"id":'AU.L2-3.3.5',"name":'Audit Correlation',"domain":'3.3 Audit and Accountability (AU)',"description":'Correlate audit record review, analysis, and reporting processes for investigation and response to indications of unlawful, unauthorized, suspicious, or unusual activity.',"objectives":[['a', 'audit record review, analysis, and reporting processes for investigation and response to indications of unlawful, unauthorized, suspicious, or unusual activity are defined; and'], ['b', 'defined audit record review, analysis, and reporting processes are correlated']]},
    {"id":'AU.L2-3.3.6',"name":'Reduction & Reporting',"domain":'3.3 Audit and Accountability (AU)',"description":'Provide audit record reduction and report generation to support on-demand analysis and reporting.',"objectives":[['a', 'an audit record reduction capability that supports on-demand analysis is provided; and'], ['b', 'a report generation capability that supports on-demand reporting is provided']]},
    {"id":'AU.L2-3.3.7',"name":'Authoritative Time Source',"domain":'3.3 Audit and Accountability (AU)',"description":'Provide a system capability that compares and synchronizes internal system clocks with an authoritative source.',"objectives":[['a', 'internal system clocks are used to generate time stamps for audit records'], ['b', 'an authoritative source with which to compare and synchronize internal system clocks is specified; and'], ['c', 'internal system clocks used to generate time stamps for audit records are compared to and synchronized with the specified authoritative time source']]},
    {"id":'AU.L2-3.3.8',"name":'Audit Protection',"domain":'3.3 Audit and Accountability (AU)',"description":'Protect audit information and audit logging tools from unauthorized access, modification, and deletion.',"objectives":[['a', 'audit information is protected from unauthorized access'], ['b', 'audit information is protected from unauthorized modification'], ['c', 'audit information is protected from unauthorized deletion'], ['d', 'audit logging tools are protected from unauthorized access'], ['e', 'audit logging tools are protected from unauthorized modification; and'], ['f', 'audit logging tools are protected from unauthorized deletion']]},
    {"id":'AU.L2-3.3.9',"name":'Audit Management',"domain":'3.3 Audit and Accountability (AU)',"description":'Limit management of audit logging functionality to a subset of privileged users.',"objectives":[['a', 'a subset of privileged users granted access to manage audit logging functionality is defined; and'], ['b', 'management of audit logging functionality is limited to the defined subset of privileged users']]},
    # Configuration Management
    {"id":'CM.L2-3.4.1',"name":'System Baselining',"domain":'3.4 Configuration Management (CM)',"description":'Establish and maintain baseline configurations and inventories of organizational systems throughout the system development life cycle.',"objectives":[['a', 'a baseline configuration is established'], ['b', 'the baseline configuration includes hardware, software, firmware, and documentation'], ['c', 'the baseline configuration is maintained (reviewed and updated) throughout the system development life cycle'], ['d', 'a system inventory is established'], ['e', 'the system inventory includes hardware, software, firmware, and documentation; and'], ['f', 'the inventory is maintained (reviewed and updated) throughout the system development life cycle']]},
    {"id":'CM.L2-3.4.2',"name":'Security Configuration Enforcement',"domain":'3.4 Configuration Management (CM)',"description":'Establish and enforce security configuration settings for information technology products employed in organizational systems.',"objectives":[['a', 'security configuration settings for information technology products employed in the system are established and included in the baseline configuration; and'], ['b', 'security configuration settings for information technology products employed in the system are enforced']]},
    {"id":'CM.L2-3.4.3',"name":'System Change Management',"domain":'3.4 Configuration Management (CM)',"description":'Track, review, approve or disapprove, and log changes to organizational systems.',"objectives":[['a', 'changes to the system are tracked'], ['b', 'changes to the system are reviewed'], ['c', 'changes to the system are approved or disapproved; and'], ['d', 'changes to the system are logged']]},
    {"id":'CM.L2-3.4.4',"name":'Security Impact Analysis',"domain":'3.4 Configuration Management (CM)',"description":'Analyze the security impact of changes prior to implementation.',"objectives":[['a', 'the security impact of changes to the system is analyzed prior to implementation']]},
    {"id":'CM.L2-3.4.5',"name":'Access Restrictions for Change',"domain":'3.4 Configuration Management (CM)',"description":'Define, document, approve, and enforce physical and logical access restrictions associated with changes to organizational systems.',"objectives":[['a', 'physical access restrictions associated with changes to the system are defined'], ['b', 'physical access restrictions associated with changes to the system are documented'], ['c', 'physical access restrictions associated with changes to the system are approved'], ['d', 'physical access restrictions associated with changes to the system are enforced'], ['e', 'logical access restrictions associated with changes to the system are defined'], ['f', 'logical access restrictions associated with changes to the system are documented'], ['g', 'logical access restrictions associated with changes to the system are approved; and'], ['h', 'logical access restrictions associated with changes to the system are enforced']]},
    {"id":'CM.L2-3.4.6',"name":'Least Functionality',"domain":'3.4 Configuration Management (CM)',"description":'Employ the principle of least functionality by configuring organizational systems to provide only essential capabilities.',"objectives":[['a', 'essential system capabilities are defined based on the principle of least functionality; and'], ['b', 'the system is configured to provide only the defined essential capabilities']]},
    {"id":'CM.L2-3.4.7',"name":'Nonessential Functionality',"domain":'3.4 Configuration Management (CM)',"description":'Restrict, disable, or prevent the use of nonessential programs, functions, ports, protocols, and services.',"objectives":[['a', 'essential programs are defined'], ['b', 'the use of nonessential programs is defined'], ['c', 'the use of nonessential programs is restricted, disabled, or prevented as defined'], ['d', 'essential functions are defined'], ['e', 'the use of nonessential functions is defined'], ['f', 'the use of nonessential functions is restricted, disabled, or prevented as defined'], ['g', 'essential ports are defined'], ['h', 'the use of nonessential ports is defined'], ['i', 'the use of nonessential ports is restricted, disabled, or prevented as defined'], ['j', 'essential protocols are defined'], ['k', 'the use of nonessential protocols is defined'], ['l', 'the use of nonessential protocols is restricted, disabled, or prevented as defined'], ['m', 'essential services are defined'], ['n', 'the use of nonessential services is defined; and'], ['o', 'the use of nonessential services is restricted, disabled, or prevented as defined']]},
    {"id":'CM.L2-3.4.8',"name":'Application Execution Policy',"domain":'3.4 Configuration Management (CM)',"description":'Apply deny-by-exception (blacklisting) policy to prevent use of unauthorized software or deny-all, permit-by-exception (whitelisting) policy to allow execution of authorized software.',"objectives":[['a', 'a policy specifying whether whitelisting or blacklisting is to be implemented is specified'], ['b', 'the software allowed to execute under whitelisting or denied use under blacklisting is specified; and'], ['c', 'whitelisting to allow the execution of authorized software or blacklisting to prevent the use of unauthorized software is implemented as specified']]},
    {"id":'CM.L2-3.4.9',"name":'User-Installed Software',"domain":'3.4 Configuration Management (CM)',"description":'Control and monitor user-installed software.',"objectives":[['a', 'a policy for controlling the installation of software by users is established'], ['b', 'installation of software by users is controlled based on the established policy; and'], ['c', 'installation of software by users is monitored']]},
    # Identification & Authentication
    {"id":'IA.L2-3.5.1',"name":'Identification',"domain":'3.5 Identification and Authentication (IA)',"description":'Identify system users, processes acting on behalf of users, and devices.',"objectives":[['a', 'system users are identified'], ['b', 'processes acting on behalf of users are identified; and'], ['c', 'devices accessing the system are identified']]},
    {"id":'IA.L2-3.5.2',"name":'Authentication',"domain":'3.5 Identification and Authentication (IA)',"description":'Authenticate (or verify) the identities of users, processes, or devices, as a prerequisite to allowing access to organizational systems.',"objectives":[['a', 'the identity of each user is authenticated or verified as a prerequisite to system access'], ['b', 'the identity of each process acting on behalf of a user is authenticated or verified as a prerequisite to system access; and'], ['c', 'the identity of each device accessing or connecting to the system is authenticated or verified as a prerequisite to system access']]},
    {"id":'IA.L2-3.5.3',"name":'Multifactor Authentication',"domain":'3.5 Identification and Authentication (IA)',"description":'Use multifactor authentication for local and network access to privileged accounts and for network access to non-privileged accounts.',"objectives":[['a', 'privileged accounts are identified'], ['b', 'multifactor authentication is implemented for local access to privileged accounts'], ['c', 'multifactor authentication is implemented for network access to privileged accounts; and'], ['d', 'multifactor authentication is implemented for network access to non-privileged accounts']]},
    {"id":'IA.L2-3.5.4',"name":'Replay-Resistant Authentication',"domain":'3.5 Identification and Authentication (IA)',"description":'Employ replay-resistant authentication mechanisms for network access to privileged and non-privileged accounts.',"objectives":[['a', 'replay-resistant authentication mechanisms are implemented for network account access to privileged and non-privileged accounts']]},
    {"id":'IA.L2-3.5.5',"name":'Identifier Reuse',"domain":'3.5 Identification and Authentication (IA)',"description":'Prevent reuse of identifiers for a defined period.',"objectives":[['a', 'a period within which identifiers cannot be reused is defined; and'], ['b', 'reuse of identifiers is prevented within the defined period']]},
    {"id":'IA.L2-3.5.6',"name":'Identifier Handling',"domain":'3.5 Identification and Authentication (IA)',"description":'Disable identifiers after a defined period of inactivity.',"objectives":[['a', 'a period of inactivity after which an identifier is disabled is defined; and'], ['b', 'identifiers are disabled after the defined period of inactivity']]},
    {"id":'IA.L2-3.5.7',"name":'Password Complexity',"domain":'3.5 Identification and Authentication (IA)',"description":'Enforce a minimum password complexity and change of characters when new passwords are created.',"objectives":[['a', 'password complexity requirements are defined'], ['b', 'password change of character requirements are defined'], ['c', 'minimum password complexity requirements as defined are enforced when new passwords are created; and'], ['d', 'minimum password change of character requirements as defined are enforced when new passwords are created']]},
    {"id":'IA.L2-3.5.8',"name":'Password Reuse',"domain":'3.5 Identification and Authentication (IA)',"description":'Prohibit password reuse for a specified number of generations.',"objectives":[['a', 'the number of generations during which a password cannot be reused is specified and'], ['b', 'reuse of passwords is prohibited during the specified number of generations']]},
    {"id":'IA.L2-3.5.9',"name":'Temporary Passwords',"domain":'3.5 Identification and Authentication (IA)',"description":'Allow temporary password use for system logons with an immediate change to a permanent password.',"objectives":[['a', 'an immediate change to a permanent password is required when a temporary password is used for system logon']]},
    {"id":'IA.L2-3.5.10',"name":'Cryptographically-Protected Passwords',"domain":'3.5 Identification and Authentication (IA)',"description":'Store and transmit only cryptographically-protected passwords.',"objectives":[['a', 'passwords are cryptographically protected in storage; and'], ['b', 'passwords are cryptographically protected in transit']]},
    {"id":'IA.L2-3.5.11',"name":'Obscure Feedback',"domain":'3.5 Identification and Authentication (IA)',"description":'Obscure feedback of authentication information.',"objectives":[['a', 'authentication information is obscured during the authentication process']]},
    # Incident Response
    {"id":'IR.L2-3.6.1',"name":'Incident Handling',"domain":'3.6 Incident Response (IR)',"description":'Establish an operational incident-handling capability for organizational systems that includes preparation, detection, analysis, containment, recovery, and user response activities.',"objectives":[['a', 'an operational incident-handling capability is established'], ['b', 'the operational incident-handling capability includes preparation'], ['c', 'the operational incident-handling capability includes detection'], ['d', 'the operational incident-handling capability includes analysis'], ['e', 'the operational incident-handling capability includes containment'], ['f', 'the operational incident-handling capability includes recovery; and'], ['g', 'the operational incident-handling capability includes user response activities']]},
    {"id":'IR.L2-3.6.2',"name":'Incident Reporting',"domain":'3.6 Incident Response (IR)',"description":'Track, document, and report incidents to designated officials and/or authorities both internal and external to the organization.',"objectives":[['a', 'incidents are tracked'], ['b', 'incidents are documented'], ['c', 'authorities to whom incidents are to be reported are identified'], ['d', 'organizational officials to whom incidents are to be reported are identified'], ['e', 'identified authorities are notified of incidents; and'], ['f', 'identified organizational officials are notified of incidents']]},
    {"id":'IR.L2-3.6.3',"name":'Incident Response Testing',"domain":'3.6 Incident Response (IR)',"description":'Test the organizational incident response capability.',"objectives":[['a', 'the incident response capability is tested']]},
    # Maintenance
    {"id":'MA.L2-3.7.1',"name":'Manage Maintenance',"domain":'3.7 Maintenance (MA)',"description":'Perform maintenance on organizational systems.',"objectives":[['a', 'system maintenance is performed']]},
    {"id":'MA.L2-3.7.2',"name":'System Maintenance Control',"domain":'3.7 Maintenance (MA)',"description":'Provide controls on the tools, techniques, mechanisms, and personnel that perform system maintenance.',"objectives":[['a', 'tools used to conduct system maintenance are controlled'], ['b', 'techniques used to conduct system maintenance are controlled'], ['c', 'mechanisms used to conduct system maintenance are controlled; and'], ['d', 'personnel used to conduct system maintenance are controlled']]},
    {"id":'MA.L2-3.7.3',"name":'Equipment Sanitization',"domain":'3.7 Maintenance (MA)',"description":'Ensure equipment removed for maintenance is sanitized with respect to all CUI.',"objectives":[['a', 'equipment to be removed from organizational spaces for off-site maintenance is sanitized of any CUI']]},
    {"id":'MA.L2-3.7.4',"name":'Media Inspection',"domain":'3.7 Maintenance (MA)',"description":'Check media containing diagnostic and test programs for malicious code before the media are used in organizational systems.',"objectives":[['a', 'media containing diagnostic and test programs are checked for malicious code before being used in organizational systems that process, store, or transmit CUI']]},
    {"id":'MA.L2-3.7.5',"name":'Multifactor Authentication for Remote Maintenance',"domain":'3.7 Maintenance (MA)',"description":'Require MFA to establish nonlocal maintenance sessions via external network connections and terminate such connections when nonlocal maintenance is complete.',"objectives":[['a', 'multifactor authentication is used to establish nonlocal maintenance sessions via external network connections; and'], ['b', 'nonlocal maintenance sessions established via external network connections are terminated when nonlocal maintenance is complete']]},
    {"id":'MA.L2-3.7.6',"name":'Remote Maintenance',"domain":'3.7 Maintenance (MA)',"description":'Supervise the maintenance activities of personnel without required access authorization.',"objectives":[['a', 'maintenance personnel without required access authorization are supervised during maintenance activities']]},
    # Media Protection
    {"id":'MP.L2-3.8.1',"name":'Media Protection',"domain":'3.8 Media Protection (MP)',"description":'Protect (i.e., physically control and securely store) system media containing CUI, both paper and digital.',"objectives":[['a', 'paper media containing CUI is physically controlled'], ['b', 'digital media containing CUI is physically controlled'], ['c', 'paper media containing CUI is securely stored; and'], ['d', 'digital media containing CUI is securely stored']]},
    {"id":'MP.L2-3.8.2',"name":'Media Access',"domain":'3.8 Media Protection (MP)',"description":'Limit access to CUI on system media to authorized users.',"objectives":[['a', 'access to CUI on system media is limited to authorized users']]},
    {"id":'MP.L2-3.8.3',"name":'Media Sanitization',"domain":'3.8 Media Protection (MP)',"description":'Sanitize or destroy system media before disposal or reuse.',"objectives":[['a', 'system media containing CUI is sanitized or destroyed before disposal; and'], ['b', 'system media containing CUI is sanitized before it is released for reuse']]},
    {"id":'MP.L2-3.8.4',"name":'Media Markings',"domain":'3.8 Media Protection (MP)',"description":'Mark media with necessary CUI markings and distribution limitations.',"objectives":[['a', 'media containing CUI is marked with applicable CUI markings; and'], ['b', 'media containing CUI is marked with distribution limitations']]},
    {"id":'MP.L2-3.8.5',"name":'Media Accountability',"domain":'3.8 Media Protection (MP)',"description":'Control access to media containing CUI and maintain accountability for media during transport.',"objectives":[['a', 'access to media containing CUI is controlled; and'], ['b', 'accountability for media containing CUI is maintained during transport outside of controlled areas']]},
    {"id":'MP.L2-3.8.6',"name":'Portable Storage Encryption',"domain":'3.8 Media Protection (MP)',"description":'Implement cryptographic mechanisms to protect the confidentiality of CUI during transport unless otherwise protected by alternative physical safeguards.',"objectives":[['a', 'the confidentiality of CUI stored on digital media is protected during transport using cryptographic mechanisms or alternative physical safeguards']]},
    {"id":'MP.L2-3.8.7',"name":'Removable Media',"domain":'3.8 Media Protection (MP)',"description":'Control the use of removable media on system components.',"objectives":[['a', 'the use of removable media on system components is controlled']]},
    {"id":'MP.L2-3.8.8',"name":'Shared Media',"domain":'3.8 Media Protection (MP)',"description":'Prohibit the use of portable storage devices when such devices have no identifiable owner.',"objectives":[['a', 'the use of portable storage devices is prohibited when such devices have no identifiable owner']]},
    {"id":'MP.L2-3.8.9',"name":'Protect CUI at Alternate Sites',"domain":'3.8 Media Protection (MP)',"description":'Protect the confidentiality of backup CUI at storage locations.',"objectives":[['a', 'the confidentiality of backup CUI is protected at storage locations']]},
    # Personnel Security
    {"id":'PS.L2-3.9.1',"name":'Screen Individuals',"domain":'3.9 Personnel Security (PS)',"description":'Screen individuals prior to authorizing access to organizational systems containing CUI.',"objectives":[['a', 'individuals are screened prior to authorizing access to organizational systems containing CUI']]},
    {"id":'PS.L2-3.9.2',"name":'Termination & Transfer',"domain":'3.9 Personnel Security (PS)',"description":'Ensure that organizational systems are protected during and after personnel actions such as terminations and transfers.',"objectives":[['a', 'a policy and/or process for terminating system access and any credentials coincident with personnel actions is established'], ['b', 'system access and credentials are terminated consistent with personnel actions such as termination or transfer; and'], ['c', 'the system is protected during and after personnel transfer actions']]},
    # Physical Protection
    {"id":'PE.L2-3.10.1',"name":'Limit Physical Access',"domain":'3.10 Physical Protection (PE)',"description":'Limit physical access to organizational systems, equipment, and the respective operating environments to authorized individuals.',"objectives":[['a', 'authorized individuals allowed physical access are identified'], ['b', 'physical access to organizational systems is limited to authorized individuals'], ['c', 'physical access to equipment is limited to authorized individuals; and'], ['d', 'physical access to operating environments is limited to authorized individuals']]},
    {"id":'PE.L2-3.10.2',"name":'Monitor Physical Access',"domain":'3.10 Physical Protection (PE)',"description":'Protect and monitor the physical facility and support infrastructure for organizational systems.',"objectives":[['a', 'the physical facility where organizational systems reside is protected'], ['b', 'the support infrastructure for organizational systems is protected'], ['c', 'the physical facility where organizational systems reside is monitored; and'], ['d', 'the support infrastructure for organizational systems is monitored']]},
    {"id":'PE.L2-3.10.3',"name":'Escort Visitors',"domain":'3.10 Physical Protection (PE)',"description":'Escort visitors and monitor visitor activity.',"objectives":[['a', 'visitors are escorted; and'], ['b', 'visitor activity is monitored']]},
    {"id":'PE.L2-3.10.4',"name":'Physical Access Logs',"domain":'3.10 Physical Protection (PE)',"description":'Maintain audit logs of physical access.',"objectives":[['a', 'audit logs of physical access are maintained']]},
    {"id":'PE.L2-3.10.5',"name":'Manage Physical Access Devices',"domain":'3.10 Physical Protection (PE)',"description":'Control and manage physical access devices.',"objectives":[['a', 'physical access devices are identified'], ['b', 'physical access devices are controlled; and'], ['c', 'physical access devices are managed']]},
    {"id":'PE.L2-3.10.6',"name":'Enforce Access Control at Alternate Work Sites',"domain":'3.10 Physical Protection (PE)',"description":'Enforce safeguarding measures for CUI at alternate work sites.',"objectives":[['a', 'safeguarding measures for CUI are defined for alternate work sites; and'], ['b', 'safeguarding measures for CUI are enforced for alternate work sites']]},
    # Risk Assessment
    {"id":'RA.L2-3.11.1',"name":'Risk Assessment',"domain":'3.11 Risk Assessment (RA)',"description":'Periodically assess the risk to organizational operations, assets, and individuals.',"objectives":[['a', 'the frequency to assess risk to organizational operations, organizational assets, and individuals is defined; and'], ['b', 'risk to organizational operations, organizational assets, and individuals resulting from the operation of an organizational system that processes, stores, or transmits CUI is assessed with the defined frequency']]},
    {"id":'RA.L2-3.11.2',"name":'Vulnerability Scan',"domain":'3.11 Risk Assessment (RA)',"description":'Scan for vulnerabilities in organizational systems and applications periodically and when new vulnerabilities affecting those systems and applications are identified.',"objectives":[['a', 'the frequency to scan for vulnerabilities in organizational systems and applications is defined'], ['b', 'vulnerability scans are performed on organizational systems with the defined frequency'], ['c', 'vulnerability scans are performed on applications with the defined frequency'], ['d', 'vulnerability scans are performed on organizational systems when new vulnerabilities are identified; and'], ['e', 'vulnerability scans are performed on applications when new vulnerabilities are identified']]},
    {"id":'RA.L2-3.11.3',"name":'Vulnerability Remediation',"domain":'3.11 Risk Assessment (RA)',"description":'Remediate vulnerabilities in accordance with risk assessments.',"objectives":[['a', 'vulnerabilities are identified; and'], ['b', 'vulnerabilities are remediated in accordance with risk assessments']]},
    # Security Assessment
    {"id":'CA.L2-3.12.1',"name":'Security Control Assessment',"domain":'3.12 Security Assessment (CA)',"description":'Periodically assess the security controls in organizational systems to determine if the controls are effective in their application.',"objectives":[['a', 'the frequency of security control assessments is defined; and'], ['b', 'security controls are assessed with the defined frequency to determine if the controls are effective in their application']]},
    {"id":'CA.L2-3.12.2',"name":'Plan of Action',"domain":'3.12 Security Assessment (CA)',"description":'Develop and implement plans of action designed to correct deficiencies and reduce or eliminate vulnerabilities in organizational systems.',"objectives":[['a', 'deficiencies and vulnerabilities to be addressed by the plan of action are identified'], ['b', 'a plan of action is developed to correct identified deficiencies and reduce or eliminate identified vulnerabilities; and'], ['c', 'the plan of action is implemented to correct identified deficiencies and reduce or eliminate identified vulnerabilities']]},
    {"id":'CA.L2-3.12.3',"name":'Security Control Monitoring',"domain":'3.12 Security Assessment (CA)',"description":'Monitor security controls on an ongoing basis to ensure the continued effectiveness of the controls.',"objectives":[['a', 'security controls are monitored on an ongoing basis to ensure the continued effectiveness of those controls']]},
    {"id":'CA.L2-3.12.4',"name":'System Security Plan',"domain":'3.12 Security Assessment (CA)',"description":'Develop, document, and periodically update system security plans that describe system boundaries, system environments of operation, how security requirements are implemented, and the relationships with or connections to other systems.',"objectives":[['a', 'a system security plan is developed'], ['b', 'the system boundary is described and documented in the system security plan'], ['c', 'the system environment of operation is described and documented in the system security plan'], ['d', 'the security requirements identified and approved by the designated authority as non-applicable are identified'], ['e', 'the method of security requirement implementation is described and documented in the system security plan'], ['f', 'the relationship with or connection to other systems is described and documented in the system security plan'], ['g', 'the frequency to update the system security plan is defined; and'], ['h', 'system security plan is updated with the defined frequency']]},
    # System & Communications Protection
    {"id":'SC.L2-3.13.1',"name":'Boundary Protection',"domain":'3.13 System and Communications Protection (SC)',"description":'Monitor, control, and protect communications at the external boundaries and key internal boundaries of organizational systems.',"objectives":[['a', 'the external system boundary is defined'], ['b', 'key internal system boundaries are defined'], ['c', 'communications are monitored at the external system boundary'], ['d', 'communications are monitored at key internal boundaries'], ['e', 'communications are controlled at the external system boundary'], ['f', 'communications are controlled at key internal boundaries'], ['g', 'communications are protected at the external system boundary; and'], ['h', 'communications are protected at key internal boundaries']]},
    {"id":'SC.L2-3.13.2',"name":'Security Engineering',"domain":'3.13 System and Communications Protection (SC)',"description":'Employ architectural designs, software development techniques, and systems engineering principles that promote effective information security within organizational systems.',"objectives":[['a', 'architectural designs that promote effective information security are identified'], ['b', 'software development techniques that promote effective information security are identified'], ['c', 'systems engineering principles that promote effective information security are identified'], ['d', 'identified architectural designs that promote effective information security are employed'], ['e', 'identified software development techniques that promote effective information security are employed; and'], ['f', 'identified systems engineering principles that promote effective information security are employed']]},
    {"id":'SC.L2-3.13.3',"name":'Role Separation',"domain":'3.13 System and Communications Protection (SC)',"description":'Separate user functionality from system management functionality.',"objectives":[['a', 'user functionality is identified'], ['b', 'system management functionality is identified; and'], ['c', 'user functionality is separated from system management functionality']]},
    {"id":'SC.L2-3.13.4',"name":'Shared Resource Control',"domain":'3.13 System and Communications Protection (SC)',"description":'Prevent unauthorized and unintended information transfer via shared system resources.',"objectives":[['a', 'unauthorized and unintended information transfer via shared system resources is prevented']]},
    {"id":'SC.L2-3.13.5',"name":'Public-Access System Separation',"domain":'3.13 System and Communications Protection (SC)',"description":'Implement subnetworks for publicly accessible system components that are physically or logically separated from internal networks.',"objectives":[['a', 'publicly accessible system components are identified; and'], ['b', 'subnetworks for publicly accessible system components are physically or logically separated from internal networks']]},
    {"id":'SC.L2-3.13.6',"name":'Network Communication by Exception',"domain":'3.13 System and Communications Protection (SC)',"description":'Deny network communications traffic by default and allow network communications traffic by exception (i.e., deny all, permit by exception).',"objectives":[['a', 'network communications traffic is denied by default; and'], ['b', 'network communications traffic is allowed by exception']]},
    {"id":'SC.L2-3.13.7',"name":'Split Tunneling',"domain":'3.13 System and Communications Protection (SC)',"description":'Prevent remote devices from simultaneously establishing non-remote connections with the system and communicating via some other connection to resources in external networks (i.e., split tunneling).',"objectives":[['a', 'remote devices are prevented from simultaneously establishing non-remote connections with the system and communicating via some other connection to resources in external networks (i.e., split tunneling)']]},
    {"id":'SC.L2-3.13.8',"name":'Data in Transit',"domain":'3.13 System and Communications Protection (SC)',"description":'Implement cryptographic mechanisms to prevent unauthorized disclosure of CUI during transmission unless otherwise protected by alternative physical safeguards.',"objectives":[['a', 'cryptographic mechanisms intended to prevent unauthorized disclosure of CUI are identified'], ['b', 'alternative physical safeguards intended to prevent unauthorized disclosure of CUI are identified; and'], ['c', 'either cryptographic mechanisms or alternative physical safeguards are implemented to prevent unauthorized disclosure of CUI during transmission']]},
    {"id":'SC.L2-3.13.9',"name":'Connections Termination',"domain":'3.13 System and Communications Protection (SC)',"description":'Terminate network connections associated with communications sessions after a defined period of inactivity.',"objectives":[['a', 'a period of inactivity to terminate network connections associated with communications sessions is defined'], ['b', 'network connections associated with communications sessions are terminated at the end of the sessions; and'], ['c', 'network connections associated with communications sessions are terminated after the defined period of inactivity']]},
    {"id":'SC.L2-3.13.10',"name":'Key Management',"domain":'3.13 System and Communications Protection (SC)',"description":'Establish and manage cryptographic keys for required cryptography employed in organizational systems.',"objectives":[['a', 'cryptographic keys are established whenever cryptography is employed; and'], ['b', 'cryptographic keys are managed whenever cryptography is employed']]},
    {"id":'SC.L2-3.13.11',"name":'CUI Encryption',"domain":'3.13 System and Communications Protection (SC)',"description":'Employ FIPS-validated cryptography when used to protect the confidentiality of CUI.',"objectives":[['a', 'FIPS-validated cryptography is employed to protect the confidentiality of CUI']]},
    {"id":'SC.L2-3.13.12',"name":'Collaborative Computing',"domain":'3.13 System and Communications Protection (SC)',"description":'Prohibit remote activation of collaborative computing devices and provide indication of use to users present at the device.',"objectives":[['a', 'collaborative computing devices are identified'], ['b', 'collaborative computing devices provide indication to users of devices in use; and'], ['c', 'remote activation of collaborative computing devices is prohibited']]},
    {"id":'SC.L2-3.13.13',"name":'Mobile Code',"domain":'3.13 System and Communications Protection (SC)',"description":'Control and monitor the use of mobile code.',"objectives":[['a', 'use of mobile code is controlled; and'], ['b', 'use of mobile code is monitored']]},
    {"id":'SC.L2-3.13.14',"name":'VoIP',"domain":'3.13 System and Communications Protection (SC)',"description":'Control and monitor the use of VoIP technologies.',"objectives":[['a', 'use of Voice over Internet Protocol (VoIP) technologies is controlled; and'], ['b', 'use of Voice over Internet Protocol (VoIP) technologies is monitored']]},
    {"id":'SC.L2-3.13.15',"name":'Communications Authenticity',"domain":'3.13 System and Communications Protection (SC)',"description":'Protect the authenticity of communications sessions.',"objectives":[['a', 'the authenticity of communications sessions is protected']]},
    {"id":'SC.L2-3.13.16',"name":'Data at Rest',"domain":'3.13 System and Communications Protection (SC)',"description":'Protect the confidentiality of CUI at rest.',"objectives":[['a', 'the confidentiality of CUI at rest is protected']]},
    # System & Information Integrity
    {"id":'SI.L2-3.14.1',"name":'Flaw Remediation',"domain":'3.14 System and Information Integrity (SI)',"description":'Identify, report, and correct information and information system flaws in a timely manner.',"objectives":[['a', 'the time within which to identify system flaws is specified'], ['b', 'system flaws are identified within the specified time frame'], ['c', 'the time within which to report system flaws is specified'], ['d', 'system flaws are reported within the specified time frame'], ['e', 'the time within which to correct system flaws is specified; and'], ['f', 'system flaws are corrected within the specified time frame']]},
    {"id":'SI.L2-3.14.2',"name":'Malicious Code Protection',"domain":'3.14 System and Information Integrity (SI)',"description":'Provide protection from malicious code at appropriate locations within organizational systems.',"objectives":[['a', 'designated locations for malicious code protection are identified; and'], ['b', 'protection from malicious code at designated locations is provided']]},
    {"id":'SI.L2-3.14.3',"name":'Security Alerts & Advisories',"domain":'3.14 System and Information Integrity (SI)',"description":'Monitor system security alerts and advisories and take action in response.',"objectives":[['a', 'response actions to system security alerts and advisories are identified'], ['b', 'system security alerts and advisories are monitored; and'], ['c', 'actions in response to system security alerts and advisories are taken']]},
    {"id":'SI.L2-3.14.4',"name":'Update Malicious Code Protection',"domain":'3.14 System and Information Integrity (SI)',"description":'Update malicious code protection mechanisms when new releases are available.',"objectives":[['a', 'malicious code protection mechanisms are updated when new releases are available']]},
    {"id":'SI.L2-3.14.5',"name":'System & File Scanning',"domain":'3.14 System and Information Integrity (SI)',"description":'Perform periodic scans of organizational systems and real-time scans of files from external sources as files are downloaded, opened, or executed.',"objectives":[['a', 'the frequency for malicious code scans is defined'], ['b', 'malicious code scans are performed with the defined frequency; and'], ['c', 'real-time malicious code scans of files from external sources as files are downloaded, opened, or executed are performed']]},
    {"id":'SI.L2-3.14.6',"name":'Monitor for Attacks',"domain":'3.14 System and Information Integrity (SI)',"description":'Monitor organizational systems, including inbound and outbound communications traffic, to detect attacks and indicators of potential attacks.',"objectives":[['a', 'the system is monitored to detect attacks and indicators of potential attacks'], ['b', 'inbound communications traffic is monitored to detect attacks and indicators of potential attacks; and'], ['c', 'outbound communications traffic is monitored to detect attacks and indicators of potential attacks']]},
    {"id":'SI.L2-3.14.7',"name":'Identify Unauthorized Use',"domain":'3.14 System and Information Integrity (SI)',"description":'Identify unauthorized use of organizational systems.',"objectives":[['a', 'authorized use of the system is defined; and'], ['b', 'unauthorized use of the system is identified']]},
]
if __name__ == '__main__':
    import os as _os
    port = int(_os.environ.get('PORT', 5000))
    is_cloud = 'PORT' in _os.environ  # Render/Railway/Heroku-style platforms set this
    host = '0.0.0.0' if is_cloud else '127.0.0.1'
    print("=" * 60)
    print("  CMMC L2 SSP Drafting Tool")
    if is_cloud:
        print("  Running on a hosted demo instance")
        print("  NOTE: this environment has ephemeral storage —")
        print("  data may reset on redeploy/restart. For real use,")
        print("  run locally via start.bat / start.sh instead.")
    else:
        print("  Running locally — no data sent to cloud")
        print(f"  Open: http://127.0.0.1:{port}")
    print("=" * 60)
    app.run(debug=False, host=host, port=port)
