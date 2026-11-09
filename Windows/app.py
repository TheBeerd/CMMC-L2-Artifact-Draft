#!/usr/bin/env python3
"""
CMMC L2 SSP Drafting Tool - Local Application
Generates System Security Plans for CMMC Level 2 compliance
"""

import os
import json
import uuid
import shutil
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, session
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

app = Flask(__name__)
app.secret_key = 'cmmc-ssp-local-tool-2024'

# Storage directory (local only)
DATA_DIR = Path(__file__).parent / 'data'
DATA_DIR.mkdir(exist_ok=True)
UPLOADS_DIR = Path(__file__).parent / 'uploads'
UPLOADS_DIR.mkdir(exist_ok=True)

SSP_FILE = DATA_DIR / 'ssp_data.json'
ASSET_FILE = DATA_DIR / 'asset_data.json'

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
ASSET_REFERENCE_MAP = {
    # AC.L2-3.1.1 — Authorized Access Control
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
        return data
    return get_default_ssp()

def save_ssp(data):
    with open(SSP_FILE, 'w') as f:
        json.dump(data, f, indent=2)

DEFAULT_BLACKLIST = [
    {"id": "bl-p2p",      "name": "Peer-to-Peer Applications",           "justification": "Increases risk of bypassing security controls, distribution of malware, and data exfiltration.",       "apps": ["BitTorrent","uTorrent","qBittorrent","Vuze","Deluge","Transmission","BitComet","FrostWire","WinMX","Soulseek","SopCast","Share","Resilio","QianQian Music","QVOD","QQLive","QQDownload","Pipi Player","PPTV","PPStream","PPMate","Kazaa","KKBOX","Gnutella","Freenet","ABC","ANts P2P","Arctic Torrent","Ares","Ares Chat Join","BBCP","CTorrent","Cabos","Cacaoweb","Direct Connect","ExtraTorrent","FileGuri","FlashGet","G3 Torrent","GNUnet","GigaTribe","JXTA","MANOLITO","Octoshape","Pirate Bay","Rufus","Syncthing","Thunder Xunlei","Thunder Xunlei Kankan","TorrentSpy","Torrentz","Trin00","Tudou Speedup","VeryCD","Winny","eMule"]},
    {"id": "bl-anon",     "name": "Anonymous Network Applications",       "justification": "Bypasses monitoring capabilities.",                                                                       "apps": ["Tor Browser","I2P","Psiphon","Lantern","N2N","Java Anon Proxy","Ultrasurf","Freegate","Your Freedom","ASProxy","AnonyTun","AnonymoX","Astrill","AutoHideIP","BlueSurface","Browsec","CGIProxy","DNS TXT Records Tunneling","DNS2TCP","DNStunnel","Easy Proxy","Epic Privacy Browser","FastStunnel","Forward","FreeMyBrowser","Freedome","GNU HTTPTunnel","GTunnel","Global Socket","Glype","HA Tunnel Plus","HTTP Tunnel","Hamachi","Hans ICMP Tunnel","Hide All IP","Hide My IP","HoSProxy","Hola Unblocker","Hotspot Shield","IP Shield","Iodine","Ivacy","JonDo","L2TP","Localtunnel","Ngrok","NthLink","OpenDoor","PPTP","PacketiX","Peer2me","Ping Tunnel","Private Tunnel","Privatix","ProXPN","Proxy HTTP","Proxy Websites","ProxyRental","Puff","Puffin","SOCKS4","SOCKS5","SSTP","SecurityKiss","Shadowsocks","Skyfire","Socks2HTTP","SoftEther","Speedify","Surf Anonymous Free","TCP Over DNS","Tailscale","TinyProxy","Tor2web","TorGuard","Tunnelier","UdpTunnel","UltraHook","Ultraviolet","VNN","VirtualShield","WASEL Pro","Wi Free","WireGuard","ZPN Connect","ZenMate","ZeroTierOne","icmptunnel"]},
    {"id": "bl-remote",   "name": "Unapproved Remote Access Applications","justification": "Allows for backdoors for threat actors and increases likelihood of lateral movement.",                  "apps": ["Phone Link","TeamViewer","AnyDesk","LogMeIn","Splashtop","Join.me","Bomgar","LiteManager","ScreenConnect","AweSun","Ammyy Admin","GoToMyPC","Alpemix","DameWare","FastViewer","LiveCare","Action1","Adobe Connect Meeting Remote Control","AeroAdmin","AirDroid","Amazon AWS Workspaces Health Status","Anyplace Control","Apple Remote Desktop","Back Orifice","Chrome Remote Desktop","Citrix Receiver","Citrix Netviewer","ConnectWise Automate","ConnectWise Automate Remote Control","DeskShare","Ericom AccessNow","FixMe IT","FleetDeck","GDB Remote","GoToAssist","GoToResolve","Goverlan","Guacamole","Honeywell TotalConnect","ISL Light","JSP File Browser","Jump Desktop","MS Netlogon","MSP360 Remote Assistant","Microsoft Service Controller","MoboRobo","N-able N-sight RMM","N-central","NTR Cloud","NetSupport Manager","Netop On Demand","Netop Remote Control","NoMachine","Parallels Access","Parsec","PhoneMyPc","PsExec","Pulseway","QQ Remote Control","RDM Plus","RDP","RDP over HTTPS","Radmin","Remote Manipulator System","RemotePC","RemotelyAnywhere","Remotix","Rexec","Rlogin","Rsh","Rsupport RemoteCall","Rsupport RemoteView","RustDesk","ShowMyPC","SimpleHelp","Solarwinds Take Control","Spy Agent","Sunlogin","Supremo","Synergy","Telnet","ToDesk","UltraViewer","VNC","VNC Clipboard","VNC File Download","VNC File Upload","WMI DCERPC","Windows Remote Management","Zoho Assist","Zoom Meeting Remote Control"]},
    {"id": "bl-crypto",   "name": "Crypto Mining Software",              "justification": "Impacts the available company resources and workstation performance.",                                   "apps": ["XMRig","cpuminer","claymore","PhoenixMiner","Ethminer","CGMiner","BFGMiner","Excavator","Bitcoin","Cardano","Ethereum","MinerGate","Monero","NiceHash BeamV3","zCash"]},
    {"id": "bl-creds",    "name": "Credential Theft Tools",              "justification": "May be used to steal credentials and escalate privileges.",                                              "apps": ["Mimikatz","Hashcat","LaZagne"]},
    {"id": "bl-vpn",      "name": "Non-Approved VPNs",                   "justification": "Increases risk of unmonitored data exfiltration.",                                                       "apps": ["ExpressVPN","Windscribe","Mullvad","TunnelBear","IPVanish","StrongVPN","Hola VPN","X-VPN","3X VPN","7VPN","ActMobile VPN","AdGuard VPN","Amaze VPN","AppVPN","Avira Phantom VPN","BeePass VPN","Bestline VPN","Betternet VPN","Blockless Proxy VPN","CAFE VPN","Cisco VPN Client","Cloudflare VPN","Croxy Proxy VPN","CyberGhost VPN","Daily VPN","DotVPN","DroidVPN","F1Rockets VPN","Fast VPN","Finch VPN","FishVPN","FlashVPN","Fly VPN","GOOSE VPN","Guardium VPN","Hexatech VPN","Hide.me VPN","HideMe Free VPN","HotVPN","Hotspot VPN","Hoxx VPN","Just Proxy VPN","Kaspersky VPN","Magic VPN","Melon VPN","Nord VPN","OKHTTP Library VPN","Octohide VPN","OpenVPN","Opera VPN","Ostrich VPN","Pinkmare VPN","Planet VPN","Plato VPN","Plex VPN","PotatoVPN","Priority VPN","Private Internet Access VPN","Proguard VPN","Proton VPN","Pure VPN","Rapid VPN","Rus VPN","SPL VPN","SaferVPN","Seed4Me VPN","Setup VPN","Simple VPN Pro","SkyVPN","Snap VPN","Speedy Quark VPN","Star VPN","Super VPN Proxy","SuperVPN Android","Surf VPN","SurfEasy VPN","Surfshark VPN","Tachyon VPN","Thunder VPN","ThunderBird VPN","TinyVPN","TipTop VPN","Touch VPN","Tower VPN","Tunnello VPN","Turbo VPN","Tuxler VPN","UFO VPN","Unseen Online Freedom VPN","V2Net VPN","V2VPN","VPN 360","VPN Go","VPN Master","VPN Master Fast Proxy Server","VPN Matreshka","VPN Monster","VPN Private","VPN Robot","VPN Vault","VeePN VPN","Vypr VPN","Wuma VPN","Yoga VPN","Zero VPN","biubiuVPN","iTop VPN"]},
    {"id": "bl-nonbiz",   "name": "Non-Business Applications (adjustable as needed)", "justification": "", "apps": ["Spotify","Pandora","Apple Music","Amazon Music","Netflix","Disney+","Amazon Prime Video","Crunchyroll","Facebook","TikTok","Instagram","LinkedIn","Epic Games","Steam","Battle.net","Xbox","Hulu","HBO Max","Paramount+","Peacock","Discord","Telegram","Pinterest","Reddit","Ubisoft Connect","Snapchat","Slack"]},
]

def load_assets():
    if ASSET_FILE.exists():
        with open(ASSET_FILE) as f:
            data = json.load(f)
        # Migrate: ensure newer sections exist on older saved files
        for key in ['assets', 'physical_security', 'authorized_users',
                    'server_apps', 'system_processes', 'ports_protocols']:
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
        # Migrate: ensure newer server app fields exist
        app_defaults = {'app_type': '', 'deployment_model': '', 'fedramp_approved': ''}
        for a in data.get('server_apps', []):
            for k, v in app_defaults.items():
                a.setdefault(k, v)
        return data
    import copy
    return {
        "assets": [],
        "physical_security": [],
        "authorized_users": [],
        "server_apps": [],
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
        "locations": [],
        "main_office_site": "",
        "home_office_site": "",
        "system_environment": "",
        "cui_data_flow": "",
        "physical_security_desc": "",
        "network_desc": "",
        "server_desc": "",
        "workstation_desc": "",
        "cloud_apps_desc": "",
        "controls_status": {},
        "controls_narratives": {},
        "images": {
            "logo": None,
            "network_diagram": None,
            "physical_diagram": None,
            "cui_flow_diagram": None,
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
    return render_template('index.html', ssp=ssp, asset_count=len(assets.get('assets', [])))

@app.route('/system-info', methods=['GET', 'POST'])
@app.route('/system-information', methods=['GET', 'POST'])
def system_info():
    ssp = load_ssp()
    if request.method == 'POST':
        data = request.form.to_dict()
        # Flat fields
        for field in ['org_name','org_abbr','mssp_name','prepared_by','system_name',
                      'system_type','system_categorization','system_uid','version',
                      'last_updated','end_users','privileged_users','cui_categories']:
            if field in data:
                ssp[field] = data[field]
        # Locations array — sent as multiple location[] values
        locations_raw = request.form.getlist('location[]')
        ssp['locations'] = [loc.strip() for loc in locations_raw if loc.strip()]
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
        for field in ['system_environment','cui_data_flow','physical_security_desc',
                      'network_desc','server_desc','workstation_desc','cloud_apps_desc',
                      'main_office_site','home_office_site']:
            ssp[field] = request.form.get(field, '')
        save_ssp(ssp)
        return redirect(url_for('environment') + '?saved=1')
    return render_template('environment.html', ssp=ssp)

@app.route('/images', methods=['GET', 'POST'])
def images():
    ssp = load_ssp()
    if request.method == 'POST':
        for img_key in ['logo', 'network_diagram', 'physical_diagram', 'cui_flow_diagram']:
            file = request.files.get(img_key)
            if file and file.filename:
                ext = Path(file.filename).suffix.lower()
                if ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp']:
                    fname = f"{img_key}{ext}"
                    fpath = UPLOADS_DIR / fname
                    file.save(fpath)
                    ssp['images'][img_key] = fname
        save_ssp(ssp)
        return redirect(url_for('images') + '?saved=1')
    return render_template('images.html', ssp=ssp)

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
                           ref_map=ASSET_REFERENCE_MAP)

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
VALID_SECTIONS = {'authorized_users', 'server_apps', 'system_processes', 'ports_protocols'}

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

@app.route('/api/import-xlsx', methods=['POST'])
def api_import_xlsx():
    """Import assets from an existing CMMC L2 Artifacts Tracker Excel file"""
    file = request.files.get('file')
    if not file:
        return jsonify({"error": "No file"}), 400
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file.read()), data_only=True)
        data = load_assets()
        imported = 0
        # Import Authorized Devices
        if 'Authorized Devices' in wb.sheetnames:
            ws = wb['Authorized Devices']
            header_row = None
            for row in ws.iter_rows(values_only=True):
                if row and str(row[0]).strip() == '#':
                    header_row = row
                    continue
                if header_row and row and row[0] and str(row[0]).strip() not in ['#', 'Ex.', 'NaN', '']:
                    try:
                        int(str(row[0]))
                        if any(v for v in row[1:] if v and str(v).strip() not in ['NaN', '']):
                            asset = {
                                'id': str(uuid.uuid4()),
                                'name': str(row[1] or '').strip(),
                                'type': str(row[2] or '').strip(),
                                'serial': str(row[3] or '').strip(),
                                'tracking_id': str(row[4] or '').strip(),
                                'category': 'CUI',
                                'location': '',
                                'owner': '',
                                'os': '',
                                'ip': '',
                                'notes': '',
                                'source': 'imported',
                            }
                            data['assets'].append(asset)
                            imported += 1
                    except (ValueError, TypeError):
                        pass
        save_assets(data)
        return jsonify({"success": True, "imported": imported})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

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
@app.route('/uploads/<filename>')
def serve_upload(filename):
    return send_file(UPLOADS_DIR / filename)

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
    for row_data in [('FCI','Low','Low','Low','GSA'), ('CUI','Medium','Low','Low','NARA'), ('ITAR','Medium','Low','Low','US Dept. of State')]:
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

    # System Environment
    add_heading(doc, 'System Environment', 2)
    doc.add_paragraph(ssp.get('system_environment', ''))

    # CUI Flow Diagram
    add_heading(doc, 'CUI Data Flow Environment', 2)
    cui_flow_path = UPLOADS_DIR / ssp['images'].get('cui_flow_diagram', '') if ssp['images'].get('cui_flow_diagram') else None
    if cui_flow_path and cui_flow_path.exists():
        doc.add_picture(str(cui_flow_path), width=Inches(5.5))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        cf = doc.add_paragraph('[CUI Flow Diagram – Insert diagram here]')
        cf.runs[0].italic = True
        cf.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    doc.add_paragraph()
    doc.add_paragraph(ssp.get('cui_data_flow', ''))

    # Physical Security Environment
    add_heading(doc, 'Physical Security Environment', 2)
    doc.add_paragraph(ssp.get('physical_security_desc', ''))

    # Physical Diagram placeholder/image
    phys_path = UPLOADS_DIR / ssp['images'].get('physical_diagram', '') if ssp['images'].get('physical_diagram') else None
    if phys_path and phys_path.exists():
        doc.add_picture(str(phys_path), width=Inches(5.5))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        ph = doc.add_paragraph('[Physical Diagram – Insert diagram here]')
        ph.runs[0].italic = True
        ph.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)

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

    # Networking
    add_heading(doc, 'Networking & Servers', 2)
    
    # Network Diagram
    net_path = UPLOADS_DIR / ssp['images'].get('network_diagram', '') if ssp['images'].get('network_diagram') else None
    if net_path and net_path.exists():
        doc.add_picture(str(net_path), width=Inches(5.5))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        nh = doc.add_paragraph('[Network Diagram – Insert diagram here]')
        nh.runs[0].italic = True
        nh.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    doc.add_paragraph()
    doc.add_paragraph(ssp.get('network_desc', ''))

    add_heading(doc, 'Servers', 3)
    doc.add_paragraph(ssp.get('server_desc', ''))
    
    # Server list from assets
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

    add_heading(doc, 'Workstations & Mobile Devices', 2)
    doc.add_paragraph(ssp.get('workstation_desc', ''))

    add_heading(doc, 'Cloud-based Applications', 2)
    doc.add_paragraph(ssp.get('cloud_apps_desc', ''))

    # ── System Inventory ──────────────────────────────────────
    add_heading(doc, 'System Inventory', 2)
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

            # Assessment objectives + narratives
            objectives = ctrl.get('objectives', [])
            ctrl_narratives = controls_narratives.get(ctrl_id, {})

            if objectives:
                obj_label = doc.add_paragraph()
                obj_run = obj_label.add_run('Assessment Objectives (Enforcement Criteria):')
                obj_run.bold = True
                obj_run.font.size = Pt(9)
                obj_run.font.color.rgb = RGBColor(0x2F, 0x54, 0x96)

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

                    # Narrative for this objective
                    narrative = sanitize_text(ctrl_narratives.get(letter, '').strip())
                    narr_p = doc.add_paragraph()
                    narr_p.paragraph_format.left_indent = Pt(28)
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
    doc.add_page_break()
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
    print("=" * 60)
    print("  CMMC L2 SSP Drafting Tool")
    print("  Running locally — no data sent to cloud")
    print(f"  Open: http://127.0.0.1:5000")
    print("=" * 60)
    app.run(debug=False, host='127.0.0.1', port=5000)
