"""Gap Analysis reference content — CMMC L2 Gap Analysis Full Environment.

Sourced from the client-supplied CMMC_L2_Gap_Analysis_Full_Environment.xlsx
workbook (316 assessment-objective-level rows across all 110 practices /
14 domains) and reconciled onto this app's own canonical CMMC_CONTROLS
objective set (320 objectives — the full NIST SP 800-171A Rev 2
determination-statement count; see app.py).

Reconciliation notes (320 app objectives vs. 316 workbook rows):
  * 292 objectives matched a workbook row by content (exact or near-exact
    text match), independent of the workbook's own [a]/[b]/... lettering —
    several controls (e.g. SC.L2-3.13.1, SC.L2-3.13.8) label objectives in
    a different order than this app, so matching is done on wording, not
    letter position.
  * 27 objectives had no reliable workbook match (the workbook consolidated
    or omitted them) and use a hand-authored question in the same voice as
    the workbook's own phrasing patterns, cross-checked against sibling
    objectives of the same control. These are exactly the objectives whose
    (ctrl_id, letter) is NOT present with a workbook-sourced 'source'
    marker — see GAP_QUESTIONS_HAND_AUTHORED below for the explicit list.

'applies_to' and 'evidence_source' are constant across all objectives of
a given control in the source workbook, so they are stored once per
control rather than once per objective.

Regenerate with the reconciliation script if CMMC_CONTROLS ever changes;
do not hand-edit the generated dict below without updating this docstring.
"""

GAP_QUESTIONS = {
    "AC.L2-3.1.1": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified authorized users across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization identified processes acting on behalf of authorized users across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "c": "Has the organization identified devices (and other systems) authorized to connect to the system across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "d": "Is system access actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?",
            "e": "Is system access actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?",
            "f": "Is system access actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?"
        }
    },
    "AC.L2-3.1.2": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization defined the types of transactions and functions that authorized users are permitted to execute? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Is system access actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?"
        }
    },
    "AC.L2-3.1.3": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization defined information flow control policies? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Has the organization defined methods and enforcement mechanisms for controlling the flow of CUI? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "c": "Has the organization identified designated sources and destinations (networks, individuals, devices) for CUI within the system and between interconnected systems across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "d": "Has the organization defined authorizations for controlling the flow of CUI? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "e": "How is approved authorizations for controlling the flow of CUI enforced (technical control, policy, or both) on the relevant system components, and can enforcement be demonstrated?"
        }
    },
    "AC.L2-3.1.4": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization defined the duties of individuals requiring separation? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: responsibilities for duties that require separation are assigned to separate individuals? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: access privileges that enable individuals to exercise the duties that require separation are granted to separate individuals? Provide supporting evidence across all relevant components."
        }
    },
    "AC.L2-3.1.5": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified privileged accounts across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Is access to privileged accounts formally authorized before use, and who grants/approves that authorization (including for vendor/third-party access)?",
            "c": "Has the organization identified security functions across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "d": "Is access to security functions formally authorized before use, and who grants/approves that authorization (including for vendor/third-party access)?"
        }
    },
    "AC.L2-3.1.6": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified nonsecurity functions across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Can the organization demonstrate: users are required to use non-privileged accounts or roles when accessing nonsecurity functions? Provide supporting evidence across all relevant components."
        }
    },
    "AC.L2-3.1.7": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization defined privileged functions? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Has the organization defined non-privileged users? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "c": "Is non-privileged users actively prevented/prohibited, and by what technical control (firewall rule, endpoint policy, etc.)?",
            "d": "Can the organization demonstrate: the execution of privileged functions is captured in audit logs? Provide supporting evidence across all relevant components."
        }
    },
    "AC.L2-3.1.8": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization defined the means of limiting unsuccessful logon attempts? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Is the defined means of limiting unsuccessful logon attempts actually implemented in practice, and how is this enforced across servers, network devices, workstations, and any third-party connections?"
        }
    },
    "AC.L2-3.1.9": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified privacy and security notices required by CUI-specified rules across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Can the organization demonstrate: privacy and security notices are displayed? Provide supporting evidence across all relevant components."
        }
    },
    "AC.L2-3.1.10": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization defined the period of inactivity after which the system initiates a session lock? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Is access to the system and viewing of data actively prevented/prohibited, and by what technical control (firewall rule, endpoint policy, etc.)?",
            "c": "Can the organization demonstrate: previously visible information is concealed via a pattern-hiding display after the defined period of inactivity? Provide supporting evidence across all relevant components."
        }
    },
    "AC.L2-3.1.11": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization defined conditions requiring a user session to terminate? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: a user session is automatically terminated after any of the defined conditions occur? Provide supporting evidence across all relevant components."
        }
    },
    "AC.L2-3.1.12": {
        "applies_to": "Remote access infrastructure (VPN concentrators, firewalls, jump servers)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Can the organization demonstrate: remote access sessions are permitted? Provide supporting evidence across all relevant components.",
            "b": "Has the organization identified the types of permitted remote access across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "c": "How is remote access sessions controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "d": "Is remote access sessions monitored? What tool or process performs this monitoring (e.g., SIEM, network monitoring, EDR) and where are results reviewed?"
        }
    },
    "AC.L2-3.1.13": {
        "applies_to": "Remote access infrastructure (VPN concentrators, firewalls, jump servers)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified cryptographic mechanisms to protect confidentiality of remote access sessions across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Is cryptographic mechanisms to protect confidentiality of remote access sessions actually implemented in practice, and how is this enforced across servers, network devices, workstations, and any third-party connections?"
        }
    },
    "AC.L2-3.1.14": {
        "applies_to": "Remote access infrastructure (VPN concentrators, firewalls, jump servers)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified managed access control points across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Can the organization demonstrate: remote access is routed through managed network access control points? Provide supporting evidence across all relevant components."
        }
    },
    "AC.L2-3.1.15": {
        "applies_to": "Remote access infrastructure (VPN concentrators, firewalls, jump servers)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified privileged commands authorized for remote execution across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization identified security-relevant information authorized to be accessed remotely across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "c": "Is the execution of the identified privileged commands via remote access formally authorized before use, and who grants/approves that authorization (including for vendor/third-party access)?",
            "d": "Is access to the identified security-relevant information via remote access formally authorized before use, and who grants/approves that authorization (including for vendor/third-party access)?"
        }
    },
    "AC.L2-3.1.16": {
        "applies_to": "Wireless network devices (access points, wireless controllers)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified wireless access points across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Is wireless access formally authorized before use, and who grants/approves that authorization (including for vendor/third-party access)?"
        }
    },
    "AC.L2-3.1.17": {
        "applies_to": "Wireless network devices (access points, wireless controllers)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "How is wireless access to the system protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system).",
            "b": "How is wireless access to the system protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system)."
        }
    },
    "AC.L2-3.1.18": {
        "applies_to": "Mobile devices",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified mobile devices that process, store, or transmit CUI across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Is mobile device connections formally authorized before use, and who grants/approves that authorization (including for vendor/third-party access)?",
            "c": "Is mobile device connections monitored? What tool or process performs this monitoring (e.g., SIEM, network monitoring, EDR) and where are results reviewed?"
        }
    },
    "AC.L2-3.1.19": {
        "applies_to": "Mobile devices",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified mobile devices and mobile computing platforms that process, store, or transmit CUI across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Can the organization demonstrate: encryption is employed to protect CUI on identified mobile devices and mobile computing platforms? Provide supporting evidence across all relevant components."
        }
    },
    "AC.L2-3.1.20": {
        "applies_to": "Third-party / external system connections (vendors, MSPs, partner networks)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified connections to external systems across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization identified the use of external systems across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "c": "Can the organization demonstrate: connections to external systems are verified? Provide supporting evidence across all relevant components.",
            "d": "Can the organization demonstrate: the use of external systems is verified? Provide supporting evidence across all relevant components.",
            "e": "How is connections to external systems controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "f": "How is the use of external systems controlled, and by what mechanism or process — including for any third-party or vendor-managed components?"
        }
    },
    "AC.L2-3.1.21": {
        "applies_to": "Third-party / external system connections; removable/portable storage",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified the use of portable storage devices containing CUI on external systems across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization defined limits on the use of portable storage devices containing CUI on external systems? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "c": "Is the use of portable storage devices containing CUI on external systems actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?"
        }
    },
    "AC.L2-3.1.22": {
        "applies_to": "Third-party / external system connections (publicly accessible systems)",
        "evidence_source": "Evidence: identity provider (AD/Entra ID) users/groups, firewall/VPN access rules, server & network device local account lists, RBAC assignments, remote access authorization records, third-party/vendor access agreements.",
        "questions": {
            "a": "Has the organization identified individuals authorized to post or process information on publicly accessible systems across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization identified procedures to ensure CUI is not posted or processed on publicly accessible systems across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "c": "Can the organization demonstrate: a review process is in place prior to posting of any content to publicly accessible systems? Provide supporting evidence across all relevant components.",
            "d": "Is content on publicly accessible systems reviewed, how often, and who performs/approves the review?",
            "e": "Can the organization demonstrate: mechanisms are in place to remove and address improper posting of CUI? Provide supporting evidence across all relevant components."
        }
    },
    "AT.L2-3.2.1": {
        "applies_to": "Personnel / training program (all roles with system access)",
        "evidence_source": "Evidence: security awareness training platform records/completion logs, phishing simulation results, role-based training records for admins of servers/network devices.",
        "questions": {
            "a": "Has the organization identified security risks associated with organizational activities involving CUI across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization identified policies, standards, and procedures related to the security of the system across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "c": "Can the organization demonstrate: managers, systems administrators, and users of the system are made aware of the security risks associated with their activities? Provide supporting evidence across all relevant components.",
            "d": "Can the organization demonstrate: managers, systems administrators, and users of the system are made aware of the applicable policies, standards, and procedures related to the security of the system? Provide supporting evidence across all relevant components."
        }
    },
    "AT.L2-3.2.2": {
        "applies_to": "Personnel / training program (all roles with system access)",
        "evidence_source": "Evidence: security awareness training platform records/completion logs, phishing simulation results, role-based training records for admins of servers/network devices.",
        "questions": {
            "a": "Has the organization defined information security-related duties, roles, and responsibilities? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: information security-related duties, roles, and responsibilities are assigned to designated personnel? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: personnel are adequately trained to carry out their assigned information security-related duties, roles, and responsibilities? Provide supporting evidence across all relevant components."
        }
    },
    "AT.L2-3.2.3": {
        "applies_to": "Personnel / training program (all roles with system access)",
        "evidence_source": "Evidence: security awareness training platform records/completion logs, phishing simulation results, role-based training records for admins of servers/network devices.",
        "questions": {
            "a": "Has the organization identified potential indicators associated with insider threats across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Can the organization demonstrate: security awareness training on recognizing and reporting potential indicators of insider threat is provided to managers and employees? Provide supporting evidence across all relevant components."
        }
    },
    "AU.L2-3.3.1": {
        "applies_to": "Servers, network devices, cloud services (logging infrastructure)",
        "evidence_source": "Evidence: server and network device (firewall, switch, router) log configuration, SIEM/log aggregation platform (e.g., Sentinel, Splunk, Wazuh) settings, retention policy, log review records, time-sync (NTP) configuration.",
        "questions": {
            "a": "Has the organization specified audit logs needed (event types to be logged) to enable monitoring, analysis, investigation, and reporting of unlawful or unauthorized system activity? Where is this documented?",
            "b": "Has the organization defined the content of audit records needed to support monitoring, analysis, investigation, and reporting of unlawful or unauthorized system activity? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "c": "Can the organization demonstrate: audit records are created (generated)? Provide supporting evidence across all relevant components.",
            "d": "Can the organization demonstrate: audit records, once created, contain the defined content? Provide supporting evidence across all relevant components.",
            "e": "Has the organization defined retention requirements for audit records? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "f": "Can the organization demonstrate: audit records are retained as defined? Provide supporting evidence across all relevant components."
        }
    },
    "AU.L2-3.3.2": {
        "applies_to": "Servers, network devices, cloud services (logging infrastructure)",
        "evidence_source": "Evidence: server and network device (firewall, switch, router) log configuration, SIEM/log aggregation platform (e.g., Sentinel, Splunk, Wazuh) settings, retention policy, log review records, time-sync (NTP) configuration.",
        "questions": {
            "a": "Has the organization defined the content of the audit records needed to support the ability to uniquely trace users to their actions? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: audit records, once created, contain the defined content? Provide supporting evidence across all relevant components."
        }
    },
    "AU.L2-3.3.3": {
        "applies_to": "Servers, network devices, cloud services (logging infrastructure)",
        "evidence_source": "Evidence: server and network device (firewall, switch, router) log configuration, SIEM/log aggregation platform (e.g., Sentinel, Splunk, Wazuh) settings, retention policy, log review records, time-sync (NTP) configuration.",
        "questions": {
            "a": "Has the organization defined a process for determining when to review logged events? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Is event types being logged reviewed, how often, and who performs/approves the review?",
            "c": "Is event types being logged kept updated across servers, network devices, and workstations? What process/cadence ensures this?"
        }
    },
    "AU.L2-3.3.4": {
        "applies_to": "Servers, network devices, cloud services (logging infrastructure)",
        "evidence_source": "Evidence: server and network device (firewall, switch, router) log configuration, SIEM/log aggregation platform (e.g., Sentinel, Splunk, Wazuh) settings, retention policy, log review records, time-sync (NTP) configuration.",
        "questions": {
            "a": "Has the organization identified personnel or roles to be alerted in the event of an audit logging process failure across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization defined types of audit logging process failures for which alert will be generated? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "c": "Can the organization demonstrate: identified personnel or roles are alerted in the event of an audit logging process failure? Provide supporting evidence across all relevant components."
        }
    },
    "AU.L2-3.3.5": {
        "applies_to": "Servers, network devices, cloud services (logging infrastructure)",
        "evidence_source": "Evidence: server and network device (firewall, switch, router) log configuration, SIEM/log aggregation platform (e.g., Sentinel, Splunk, Wazuh) settings, retention policy, log review records, time-sync (NTP) configuration.",
        "questions": {
            "a": "Has the organization defined audit record review, analysis, and reporting processes for investigation and response to indications of unlawful, unauthorized, suspicious, or unusual activity? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: defined audit record review, analysis, and reporting processes are correlated? Provide supporting evidence across all relevant components."
        }
    },
    "AU.L2-3.3.6": {
        "applies_to": "Servers, network devices, cloud services (logging infrastructure)",
        "evidence_source": "Evidence: server and network device (firewall, switch, router) log configuration, SIEM/log aggregation platform (e.g., Sentinel, Splunk, Wazuh) settings, retention policy, log review records, time-sync (NTP) configuration.",
        "questions": {
            "a": "Can the organization demonstrate: an audit record reduction capability that supports on-demand analysis is provided? Provide supporting evidence across all relevant components.",
            "b": "Can the organization demonstrate: a report generation capability that supports on-demand reporting is provided? Provide supporting evidence across all relevant components."
        }
    },
    "AU.L2-3.3.7": {
        "applies_to": "Servers, network devices, cloud services (logging infrastructure)",
        "evidence_source": "Evidence: server and network device (firewall, switch, router) log configuration, SIEM/log aggregation platform (e.g., Sentinel, Splunk, Wazuh) settings, retention policy, log review records, time-sync (NTP) configuration.",
        "questions": {
            "a": "Can the organization demonstrate: internal system clocks are used to generate time stamps for audit records? Provide supporting evidence across all relevant components.",
            "b": "Has the organization specified an authoritative source with which to compare and synchronize internal system clocks? Where is this documented?",
            "c": "Can the organization demonstrate: internal system clocks used to generate time stamps for audit records are compared to and synchronized with the specified authoritative time source? Provide supporting evidence across all relevant components."
        }
    },
    "AU.L2-3.3.8": {
        "applies_to": "Servers, network devices, cloud services (logging infrastructure)",
        "evidence_source": "Evidence: server and network device (firewall, switch, router) log configuration, SIEM/log aggregation platform (e.g., Sentinel, Splunk, Wazuh) settings, retention policy, log review records, time-sync (NTP) configuration.",
        "questions": {
            "a": "How is audit information protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system).",
            "b": "How is audit information protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system).",
            "c": "How is audit information protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system).",
            "d": "How is audit logging tools protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system).",
            "e": "How is audit logging tools protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system).",
            "f": "How is audit logging tools protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system)."
        }
    },
    "AU.L2-3.3.9": {
        "applies_to": "Servers, network devices, cloud services (logging infrastructure)",
        "evidence_source": "Evidence: server and network device (firewall, switch, router) log configuration, SIEM/log aggregation platform (e.g., Sentinel, Splunk, Wazuh) settings, retention policy, log review records, time-sync (NTP) configuration.",
        "questions": {
            "a": "Has the organization defined a subset of privileged users granted access to manage audit logging functionality? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Is management of audit logging functionality actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?"
        }
    },
    "CM.L2-3.4.1": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: baseline configuration standards for servers, network devices, and workstations (e.g., CIS Benchmarks, DISA STIGs), change control records/tickets, configuration management tooling (Intune, Ansible, SCCM, vendor management consoles), inventory of hardware/software/firmware.",
        "questions": {
            "a": "Has a baseline configuration been established, and how is it maintained across servers, network devices, and the broader environment?",
            "b": "Can the organization demonstrate: the baseline configuration includes hardware, software, firmware, and documentation? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: the baseline configuration is maintained (reviewed and updated) throughout the system development life cycle? Provide supporting evidence across all relevant components.",
            "d": "Has a system inventory been established, and how is it maintained across servers, network devices, and the broader environment?",
            "e": "Can the organization demonstrate: the system inventory includes hardware, software, firmware, and documentation? Provide supporting evidence across all relevant components.",
            "f": "Can the organization demonstrate: the inventory is maintained (reviewed and updated) throughout the system development life cycle? Provide supporting evidence across all relevant components."
        }
    },
    "CM.L2-3.4.2": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: baseline configuration standards for servers, network devices, and workstations (e.g., CIS Benchmarks, DISA STIGs), change control records/tickets, configuration management tooling (Intune, Ansible, SCCM, vendor management consoles), inventory of hardware/software/firmware.",
        "questions": {
            "a": "Has security configuration settings for IT products employed in the system been established, and how is it maintained across servers, network devices, and the broader environment?",
            "b": "How is security configuration settings for IT products employed in the system enforced (technical control, policy, or both) on the relevant system components, and can enforcement be demonstrated?"
        }
    },
    "CM.L2-3.4.3": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: baseline configuration standards for servers, network devices, and workstations (e.g., CIS Benchmarks, DISA STIGs), change control records/tickets, configuration management tooling (Intune, Ansible, SCCM, vendor management consoles), inventory of hardware/software/firmware.",
        "questions": {
            "a": "Is changes to the system tracked? In what system/log?",
            "b": "Is changes to the system reviewed, how often, and who performs/approves the review?",
            "c": "Can the organization demonstrate: changes to the system are approved or disapproved? Provide supporting evidence across all relevant components.",
            "d": "Can the organization demonstrate: changes to the system are logged? Provide supporting evidence across all relevant components."
        }
    },
    "CM.L2-3.4.4": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: baseline configuration standards for servers, network devices, and workstations (e.g., CIS Benchmarks, DISA STIGs), change control records/tickets, configuration management tooling (Intune, Ansible, SCCM, vendor management consoles), inventory of hardware/software/firmware.",
        "questions": {
            "a": "Can the organization demonstrate: the security impact of changes to the system is analyzed prior to implementation? Provide supporting evidence across all relevant components."
        }
    },
    "CM.L2-3.4.5": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: baseline configuration standards for servers, network devices, and workstations (e.g., CIS Benchmarks, DISA STIGs), change control records/tickets, configuration management tooling (Intune, Ansible, SCCM, vendor management consoles), inventory of hardware/software/firmware.",
        "questions": {
            "a": "Has the organization defined physical access restrictions associated with changes to the system? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Is physical access restrictions associated with changes to the system documented? Where, and is the documentation current for all in-scope servers, network devices, and connections?",
            "c": "Can the organization demonstrate: physical access restrictions associated with changes to the system are approved? Provide supporting evidence across all relevant components.",
            "d": "How is physical access restrictions associated with changes to the system enforced (technical control, policy, or both) on the relevant system components, and can enforcement be demonstrated?",
            "e": "Has the organization defined logical access restrictions associated with changes to the system? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "f": "Is logical access restrictions associated with changes to the system documented? Where, and is the documentation current for all in-scope servers, network devices, and connections?",
            "g": "Can the organization demonstrate: logical access restrictions associated with changes to the system are approved? Provide supporting evidence across all relevant components.",
            "h": "How is logical access restrictions associated with changes to the system enforced (technical control, policy, or both) on the relevant system components, and can enforcement be demonstrated?"
        }
    },
    "CM.L2-3.4.6": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: baseline configuration standards for servers, network devices, and workstations (e.g., CIS Benchmarks, DISA STIGs), change control records/tickets, configuration management tooling (Intune, Ansible, SCCM, vendor management consoles), inventory of hardware/software/firmware.",
        "questions": {
            "a": "Has the organization defined essential system capabilities? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: the system is configured to provide only the defined essential capabilities? Provide supporting evidence across all relevant components."
        }
    },
    "CM.L2-3.4.7": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: baseline configuration standards for servers, network devices, and workstations (e.g., CIS Benchmarks, DISA STIGs), change control records/tickets, configuration management tooling (Intune, Ansible, SCCM, vendor management consoles), inventory of hardware/software/firmware.",
        "questions": {
            "a": "Has the organization defined essential programs? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Has the organization defined the use of nonessential programs? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "c": "Is the use of nonessential programs actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?",
            "d": "Has the organization defined essential functions? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "e": "Has the organization defined the use of nonessential functions? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "f": "Is the use of nonessential functions actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?",
            "g": "Has the organization defined essential ports? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "h": "Has the organization defined the use of nonessential ports? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "i": "Is the use of nonessential ports actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?",
            "j": "Has the organization defined essential protocols? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "k": "Has the organization defined the use of nonessential protocols? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "l": "Is the use of nonessential protocols actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?",
            "m": "Has the organization defined essential services? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "n": "Has the organization defined the use of nonessential services? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "o": "Is the use of nonessential services actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?"
        }
    },
    "CM.L2-3.4.8": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: baseline configuration standards for servers, network devices, and workstations (e.g., CIS Benchmarks, DISA STIGs), change control records/tickets, configuration management tooling (Intune, Ansible, SCCM, vendor management consoles), inventory of hardware/software/firmware.",
        "questions": {
            "a": "Has the organization specified a policy specifying whether whitelisting or blacklisting is to be implemented? Where is this documented?",
            "b": "Has the organization specified the software allowed to execute under whitelisting or denied use under blacklisting? Where is this documented?",
            "c": "Is whitelisting to allow execution of authorized software or blacklisting to prevent use of unauthorized software actually implemented in practice, and how is this enforced across servers, network devices, workstations, and any third-party connections?"
        }
    },
    "CM.L2-3.4.9": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: baseline configuration standards for servers, network devices, and workstations (e.g., CIS Benchmarks, DISA STIGs), change control records/tickets, configuration management tooling (Intune, Ansible, SCCM, vendor management consoles), inventory of hardware/software/firmware.",
        "questions": {
            "a": "Has a policy for controlling installation of software by users been established, and how is it maintained across servers, network devices, and the broader environment?",
            "b": "How is installation of software by users controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "c": "Is installation of software by users monitored? What tool or process performs this monitoring (e.g., SIEM, network monitoring, EDR) and where are results reviewed?"
        }
    },
    "IA.L2-3.5.1": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Has the organization identified system users across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization identified processes acting on behalf of users across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "c": "Has the organization identified devices accessing the system across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory."
        }
    },
    "IA.L2-3.5.2": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Can the organization demonstrate: the identity of each user is authenticated or verified as a prerequisite to system access? Provide supporting evidence across all relevant components.",
            "b": "Can the organization demonstrate: the identity of each process acting on behalf of a user is authenticated or verified as a prerequisite to system access? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: the identity of each device accessing or connecting to the system is authenticated or verified as a prerequisite to system access? Provide supporting evidence across all relevant components."
        }
    },
    "IA.L2-3.5.3": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Has the organization identified privileged accounts across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Is multifactor authentication actually implemented in practice, and how is this enforced across servers, network devices, workstations, and any third-party connections?",
            "c": "Is multifactor authentication actually implemented in practice, and how is this enforced across servers, network devices, workstations, and any third-party connections?",
            "d": "Is multifactor authentication actually implemented in practice, and how is this enforced across servers, network devices, workstations, and any third-party connections?"
        }
    },
    "IA.L2-3.5.4": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Is replay-resistant authentication mechanisms actually implemented in practice, and how is this enforced across servers, network devices, workstations, and any third-party connections?"
        }
    },
    "IA.L2-3.5.5": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Has the organization defined a period within which identifiers cannot be reused? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Is reuse of identifiers actively prevented/prohibited, and by what technical control (firewall rule, endpoint policy, etc.)?"
        }
    },
    "IA.L2-3.5.6": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Has the organization defined a period of inactivity after which an identifier is disabled? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: identifiers are disabled after the defined period of inactivity? Provide supporting evidence across all relevant components."
        }
    },
    "IA.L2-3.5.7": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Has the organization defined password complexity requirements? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Has the organization defined password change of character requirements? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "c": "How is minimum password complexity requirements as defined enforced (technical control, policy, or both) on the relevant system components, and can enforcement be demonstrated?",
            "d": "How is minimum password change of character requirements as defined enforced (technical control, policy, or both) on the relevant system components, and can enforcement be demonstrated?"
        }
    },
    "IA.L2-3.5.8": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Has the organization specified the number of generations during which a password cannot be reused? Where is this documented?",
            "b": "Is reuse of passwords actively prevented/prohibited, and by what technical control (firewall rule, endpoint policy, etc.)?"
        }
    },
    "IA.L2-3.5.9": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Can the organization demonstrate: an immediate change to a permanent password is required when a temporary password is used for system logon? Provide supporting evidence across all relevant components."
        }
    },
    "IA.L2-3.5.10": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Can the organization demonstrate: passwords are cryptographically protected in storage? Provide supporting evidence across all relevant components.",
            "b": "Can the organization demonstrate: passwords are cryptographically protected in transit? Provide supporting evidence across all relevant components."
        }
    },
    "IA.L2-3.5.11": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: identity provider authentication policies (AD/Entra ID), MFA configuration for servers/network device admin access, local account and service account inventories, password policy configuration across platforms.",
        "questions": {
            "a": "Can the organization demonstrate: authentication information is obscured during the authentication process? Provide supporting evidence across all relevant components."
        }
    },
    "IR.L2-3.6.1": {
        "applies_to": "Incident response program (all components)",
        "evidence_source": "Evidence: written Incident Response Plan, SOC/SIEM or EDR incident records, IR test/tabletop documentation, third-party incident notification clauses, DIBNet/DC3 reporting records.",
        "questions": {
            "a": "Has an operational incident-handling capability been established, and how is it maintained across servers, network devices, and the broader environment?",
            "b": "Can the organization demonstrate: the operational incident-handling capability includes preparation? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: the operational incident-handling capability includes detection? Provide supporting evidence across all relevant components.",
            "d": "Can the organization demonstrate: the operational incident-handling capability includes analysis? Provide supporting evidence across all relevant components.",
            "e": "Can the organization demonstrate: the operational incident-handling capability includes containment? Provide supporting evidence across all relevant components.",
            "f": "Can the organization demonstrate: the operational incident-handling capability includes recovery? Provide supporting evidence across all relevant components.",
            "g": "Can the organization demonstrate: the operational incident-handling capability includes user response activities? Provide supporting evidence across all relevant components."
        }
    },
    "IR.L2-3.6.2": {
        "applies_to": "Incident response program (all components)",
        "evidence_source": "Evidence: written Incident Response Plan, SOC/SIEM or EDR incident records, IR test/tabletop documentation, third-party incident notification clauses, DIBNet/DC3 reporting records.",
        "questions": {
            "a": "Is incidents tracked? In what system/log?",
            "b": "Is incidents documented? Where, and is the documentation current for all in-scope servers, network devices, and connections?",
            "c": "Has the organization identified authorities to whom incidents are to be reported across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "d": "Has the organization identified organizational officials to whom incidents are to be reported across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "e": "Are identified authorities actually notified in practice? Show a recent example or the notification workflow, including any third-party/vendor notification obligations.",
            "f": "Are identified organizational officials actually notified in practice? Show a recent example or the notification workflow, including any third-party/vendor notification obligations."
        }
    },
    "IR.L2-3.6.3": {
        "applies_to": "Incident response program (all components)",
        "evidence_source": "Evidence: written Incident Response Plan, SOC/SIEM or EDR incident records, IR test/tabletop documentation, third-party incident notification clauses, DIBNet/DC3 reporting records.",
        "questions": {
            "a": "Is the incident response capability tested? How often, and where are test results/after-action items recorded?"
        }
    },
    "MA.L2-3.7.1": {
        "applies_to": "Servers, network devices, workstations (all maintained hardware)",
        "evidence_source": "Evidence: maintenance schedules and records for servers and network devices (patch/firmware history), remote maintenance session logs (VPN/jump box, MFA for vendor remote access), third-party/vendor maintenance agreements and supervision records, media sanitization records for decommissioned hardware.",
        "questions": {
            "a": "Can the organization demonstrate: system maintenance is performed? Provide supporting evidence across all relevant components."
        }
    },
    "MA.L2-3.7.2": {
        "applies_to": "Servers, network devices, workstations (all maintained hardware)",
        "evidence_source": "Evidence: maintenance schedules and records for servers and network devices (patch/firmware history), remote maintenance session logs (VPN/jump box, MFA for vendor remote access), third-party/vendor maintenance agreements and supervision records, media sanitization records for decommissioned hardware.",
        "questions": {
            "a": "How is tools used to conduct system maintenance controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "b": "How is techniques used to conduct system maintenance controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "c": "How is mechanisms used to conduct system maintenance controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "d": "How is personnel used to conduct system maintenance controlled, and by what mechanism or process — including for any third-party or vendor-managed components?"
        }
    },
    "MA.L2-3.7.3": {
        "applies_to": "Servers, network devices, workstations (equipment leaving the facility)",
        "evidence_source": "Evidence: maintenance schedules and records for servers and network devices (patch/firmware history), remote maintenance session logs (VPN/jump box, MFA for vendor remote access), third-party/vendor maintenance agreements and supervision records, media sanitization records for decommissioned hardware.",
        "questions": {
            "a": "Can the organization demonstrate: equipment removed for off-site maintenance is sanitized of any CUI? Provide supporting evidence across all relevant components."
        }
    },
    "MA.L2-3.7.4": {
        "applies_to": "Servers, network devices, workstations (diagnostic/test media)",
        "evidence_source": "Evidence: maintenance schedules and records for servers and network devices (patch/firmware history), remote maintenance session logs (VPN/jump box, MFA for vendor remote access), third-party/vendor maintenance agreements and supervision records, media sanitization records for decommissioned hardware.",
        "questions": {
            "a": "Can the organization demonstrate: media containing diagnostic and test programs is checked for malicious code before the media are used in organizational systems? Provide supporting evidence across all relevant components."
        }
    },
    "MA.L2-3.7.5": {
        "applies_to": "Servers, network devices (nonlocal/vendor remote maintenance sessions)",
        "evidence_source": "Evidence: maintenance schedules and records for servers and network devices (patch/firmware history), remote maintenance session logs (VPN/jump box, MFA for vendor remote access), third-party/vendor maintenance agreements and supervision records, media sanitization records for decommissioned hardware.",
        "questions": {
            "a": "Can the organization demonstrate: multifactor authentication is used to establish nonlocal maintenance sessions via external network connections? Provide supporting evidence across all relevant components.",
            "b": "Can the organization demonstrate: nonlocal maintenance sessions established via external network connections are terminated when nonlocal maintenance is complete? Provide supporting evidence across all relevant components."
        }
    },
    "MA.L2-3.7.6": {
        "applies_to": "Servers, network devices (third-party maintenance personnel)",
        "evidence_source": "Evidence: maintenance schedules and records for servers and network devices (patch/firmware history), remote maintenance session logs (VPN/jump box, MFA for vendor remote access), third-party/vendor maintenance agreements and supervision records, media sanitization records for decommissioned hardware.",
        "questions": {
            "a": "Has the organization identified maintenance personnel without required access authorization across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory."
        }
    },
    "MP.L2-3.8.1": {
        "applies_to": "Servers, workstations, backup media (paper and digital CUI)",
        "evidence_source": "Evidence: media handling and marking procedures, removable media restriction policy (servers, network devices, workstations), encryption status of backup/removable media, backup storage location security (on-prem or cloud), media destruction/sanitization records.",
        "questions": {
            "a": "Can the organization demonstrate: paper media containing CUI is physically controlled? Provide supporting evidence across all relevant components.",
            "b": "Can the organization demonstrate: digital media containing CUI is physically controlled? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: paper media containing CUI is securely stored? Provide supporting evidence across all relevant components.",
            "d": "Can the organization demonstrate: digital media containing CUI is securely stored? Provide supporting evidence across all relevant components."
        }
    },
    "MP.L2-3.8.2": {
        "applies_to": "Servers, workstations, removable/backup media",
        "evidence_source": "Evidence: media handling and marking procedures, removable media restriction policy (servers, network devices, workstations), encryption status of backup/removable media, backup storage location security (on-prem or cloud), media destruction/sanitization records.",
        "questions": {
            "a": "Is access to CUI on system media (paper and digital) actually restricted/limited to authorized users, and what control enforces this restriction on the relevant servers, network devices, or connections?"
        }
    },
    "MP.L2-3.8.3": {
        "applies_to": "Servers, workstations, removable/backup media",
        "evidence_source": "Evidence: media handling and marking procedures, removable media restriction policy (servers, network devices, workstations), encryption status of backup/removable media, backup storage location security (on-prem or cloud), media destruction/sanitization records.",
        "questions": {
            "a": "Can the organization demonstrate: system media containing CUI is sanitized or destroyed before disposal? Provide supporting evidence across all relevant components.",
            "b": "Can the organization demonstrate: system media containing CUI is sanitized before it is released for reuse? Provide supporting evidence across all relevant components."
        }
    },
    "MP.L2-3.8.4": {
        "applies_to": "Servers, workstations, removable/backup media",
        "evidence_source": "Evidence: media handling and marking procedures, removable media restriction policy (servers, network devices, workstations), encryption status of backup/removable media, backup storage location security (on-prem or cloud), media destruction/sanitization records.",
        "questions": {
            "a": "Can the organization demonstrate: media containing CUI is marked with applicable CUI markings? Provide supporting evidence across all relevant components.",
            "b": "Can the organization demonstrate: media containing CUI is marked with distribution limitations? Provide supporting evidence across all relevant components."
        }
    },
    "MP.L2-3.8.5": {
        "applies_to": "Servers, workstations, removable/backup media in transport",
        "evidence_source": "Evidence: media handling and marking procedures, removable media restriction policy (servers, network devices, workstations), encryption status of backup/removable media, backup storage location security (on-prem or cloud), media destruction/sanitization records.",
        "questions": {
            "a": "Is access to media containing CUI actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?",
            "b": "Can the organization demonstrate: accountability for media containing CUI is maintained during transport outside of controlled areas? Provide supporting evidence across all relevant components."
        }
    },
    "MP.L2-3.8.6": {
        "applies_to": "Removable/portable storage and backup media",
        "evidence_source": "Evidence: media handling and marking procedures, removable media restriction policy (servers, network devices, workstations), encryption status of backup/removable media, backup storage location security (on-prem or cloud), media destruction/sanitization records.",
        "questions": {
            "a": "Is cryptographic mechanisms or alternative physical safeguards actually implemented in practice, and how is this enforced across servers, network devices, workstations, and any third-party connections?"
        }
    },
    "MP.L2-3.8.7": {
        "applies_to": "Servers, network devices, workstations (removable media ports)",
        "evidence_source": "Evidence: media handling and marking procedures, removable media restriction policy (servers, network devices, workstations), encryption status of backup/removable media, backup storage location security (on-prem or cloud), media destruction/sanitization records.",
        "questions": {
            "a": "How is the use of removable media on system components controlled, and by what mechanism or process — including for any third-party or vendor-managed components?"
        }
    },
    "MP.L2-3.8.8": {
        "applies_to": "Servers, workstations, removable/backup media",
        "evidence_source": "Evidence: media handling and marking procedures, removable media restriction policy (servers, network devices, workstations), encryption status of backup/removable media, backup storage location security (on-prem or cloud), media destruction/sanitization records.",
        "questions": {
            "a": "Is the use of portable storage devices actively prevented/prohibited, and by what technical control (firewall rule, endpoint policy, etc.)?"
        }
    },
    "MP.L2-3.8.9": {
        "applies_to": "Servers, workstations, removable/backup media",
        "evidence_source": "Evidence: media handling and marking procedures, removable media restriction policy (servers, network devices, workstations), encryption status of backup/removable media, backup storage location security (on-prem or cloud), media destruction/sanitization records.",
        "questions": {
            "a": "How is the confidentiality of backup CUI at storage locations protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system)."
        }
    },
    "PS.L2-3.9.1": {
        "applies_to": "Personnel / HR process",
        "evidence_source": "Evidence: HR screening/background-check procedure, onboarding/offboarding checklist, account deprovisioning records across all systems (AD/Entra ID, VPN, network device admin accounts) tied to termination or role change.",
        "questions": {
            "a": "Are individuals screened (background check) prior to being granted access to CUI, and is this documented per role/risk designation?"
        }
    },
    "PS.L2-3.9.2": {
        "applies_to": "Personnel / HR process",
        "evidence_source": "Evidence: HR screening/background-check procedure, onboarding/offboarding checklist, account deprovisioning records across all systems (AD/Entra ID, VPN, network device admin accounts) tied to termination or role change.",
        "questions": {
            "a": "Has the organization established a policy and/or process for terminating system access and any credentials coincident with personnel actions? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: system access and credentials are terminated consistent with personnel actions such as termination or transfer? Provide supporting evidence across all relevant components.",
            "c": "How is organizational systems containing CUI protected during and after personnel transfer actions? Identify the specific control on each relevant component (server, network device, workstation, or third-party system)."
        }
    },
    "PE.L2-3.10.1": {
        "applies_to": "Facility, server rooms, network/data closets",
        "evidence_source": "Evidence: physical access control records for server rooms/network closets/data centers (badge logs, visitor logs), environmental/monitoring controls, physical security of on-prem servers and network equipment, alternate work site / telework safeguards.",
        "questions": {
            "a": "Has the organization identified authorized individuals allowed physical access across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Is physical access to organizational systems actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?",
            "c": "Is physical access to equipment actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?",
            "d": "Is physical access to the respective operating environments actually restricted/limited, and what control enforces this restriction on the relevant servers, network devices, or connections?"
        }
    },
    "PE.L2-3.10.2": {
        "applies_to": "Facility, server rooms, network/data closets, support infrastructure (power, cooling)",
        "evidence_source": "Evidence: physical access control records for server rooms/network closets/data centers (badge logs, visitor logs), environmental/monitoring controls, physical security of on-prem servers and network equipment, alternate work site / telework safeguards.",
        "questions": {
            "a": "How is the physical facility where organizational systems reside protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system).",
            "b": "How is support infrastructure for organizational systems protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system).",
            "c": "How is the physical facility where organizational systems reside monitored, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "d": "How is the support infrastructure for organizational systems monitored, and by what mechanism or process — including for any third-party or vendor-managed components?"
        }
    },
    "PE.L2-3.10.3": {
        "applies_to": "Facility, server rooms, network/data closets",
        "evidence_source": "Evidence: physical access control records for server rooms/network closets/data centers (badge logs, visitor logs), environmental/monitoring controls, physical security of on-prem servers and network equipment, alternate work site / telework safeguards.",
        "questions": {
            "a": "Are visitors (including vendor/maintenance personnel) escorted while in server rooms or network closets, and is this enforced/documented?",
            "b": "Is visitor activity monitored? What tool or process performs this monitoring (e.g., SIEM, network monitoring, EDR) and where are results reviewed?"
        }
    },
    "PE.L2-3.10.4": {
        "applies_to": "Facility, server rooms, network/data closets",
        "evidence_source": "Evidence: physical access control records for server rooms/network closets/data centers (badge logs, visitor logs), environmental/monitoring controls, physical security of on-prem servers and network equipment, alternate work site / telework safeguards.",
        "questions": {
            "a": "Can the organization demonstrate: audit logs of physical access are maintained? Provide supporting evidence across all relevant components."
        }
    },
    "PE.L2-3.10.5": {
        "applies_to": "Facility physical access devices (badge readers, locks, cameras)",
        "evidence_source": "Evidence: physical access control records for server rooms/network closets/data centers (badge logs, visitor logs), environmental/monitoring controls, physical security of on-prem servers and network equipment, alternate work site / telework safeguards.",
        "questions": {
            "a": "Has the organization identified physical access devices across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "How is physical access devices controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "c": "Can the organization demonstrate: physical access devices are managed? Provide supporting evidence across all relevant components."
        }
    },
    "PE.L2-3.10.6": {
        "applies_to": "Remote/telework work sites",
        "evidence_source": "Evidence: physical access control records for server rooms/network closets/data centers (badge logs, visitor logs), environmental/monitoring controls, physical security of on-prem servers and network equipment, alternate work site / telework safeguards.",
        "questions": {
            "a": "Has the organization defined safeguarding measures for CUI? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Is safeguarding measures for CUI actually implemented in practice, and how is this enforced across servers, network devices, workstations, and any third-party connections?"
        }
    },
    "RA.L2-3.11.1": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: risk assessment schedule/report, vulnerability scan results covering servers, network devices, and workstations (e.g., Nessus, Qualys, Defender Vulnerability Management), scan frequency configuration, third-party/vendor risk assessments.",
        "questions": {
            "a": "Has the organization defined the frequency to assess risk to organizational operations, assets, and individuals? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: risk to organizational operations, assets, and individuals resulting from the operation of a system that processes, stores, or transmits CUI is assessed at the defined frequency? Provide supporting evidence across all relevant components."
        }
    },
    "RA.L2-3.11.2": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: risk assessment schedule/report, vulnerability scan results covering servers, network devices, and workstations (e.g., Nessus, Qualys, Defender Vulnerability Management), scan frequency configuration, third-party/vendor risk assessments.",
        "questions": {
            "a": "Has the organization defined the frequency to scan for vulnerabilities in organizational systems and applications? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: vulnerability scans are performed on organizational systems at the defined frequency? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: vulnerability scans are performed on applications at the defined frequency? Provide supporting evidence across all relevant components.",
            "d": "Has the organization identified vulnerability scans are performed when new vulnerabilities potentially affecting organizational systems and applications across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "e": "Has the organization identified vulnerability scans are performed on applications when new vulnerabilities are identified across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory."
        }
    },
    "RA.L2-3.11.3": {
        "applies_to": "All system components (servers, network devices, workstations, cloud services)",
        "evidence_source": "Evidence: risk assessment schedule/report, vulnerability scan results covering servers, network devices, and workstations (e.g., Nessus, Qualys, Defender Vulnerability Management), scan frequency configuration, third-party/vendor risk assessments.",
        "questions": {
            "a": "Has the organization identified vulnerabilities across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Can the organization demonstrate: vulnerabilities are remediated in accordance with risk assessments? Provide supporting evidence across all relevant components."
        }
    },
    "CA.L2-3.12.1": {
        "applies_to": "Program-level (SSP / POA&M covering full environment)",
        "evidence_source": "Evidence: System Security Plan (SSP) covering the full environment (servers, network devices, cloud services, third-party connections), Plan of Action & Milestones (POA&M), security control assessment reports, continuous monitoring dashboards.",
        "questions": {
            "a": "Has the organization defined the frequency of security control assessments? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: security controls are assessed at the defined frequency to determine if the controls are effective in their application? Provide supporting evidence across all relevant components."
        }
    },
    "CA.L2-3.12.2": {
        "applies_to": "Program-level (SSP / POA&M covering full environment)",
        "evidence_source": "Evidence: System Security Plan (SSP) covering the full environment (servers, network devices, cloud services, third-party connections), Plan of Action & Milestones (POA&M), security control assessment reports, continuous monitoring dashboards.",
        "questions": {
            "a": "Has the organization identified deficiencies and vulnerabilities to be addressed by the plan of action across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Can the organization demonstrate: a plan of action is developed to correct identified deficiencies and reduce or eliminate vulnerabilities? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: the plan of action is implemented to correct identified deficiencies and reduce or eliminate identified vulnerabilities? Provide supporting evidence across all relevant components."
        }
    },
    "CA.L2-3.12.3": {
        "applies_to": "Program-level (SSP / POA&M covering full environment)",
        "evidence_source": "Evidence: System Security Plan (SSP) covering the full environment (servers, network devices, cloud services, third-party connections), Plan of Action & Milestones (POA&M), security control assessment reports, continuous monitoring dashboards.",
        "questions": {
            "a": "Is security controls monitored? What tool or process performs this monitoring (e.g., SIEM, network monitoring, EDR) and where are results reviewed?"
        }
    },
    "CA.L2-3.12.4": {
        "applies_to": "Program-level (SSP / POA&M covering full environment)",
        "evidence_source": "Evidence: System Security Plan (SSP) covering the full environment (servers, network devices, cloud services, third-party connections), Plan of Action & Milestones (POA&M), security control assessment reports, continuous monitoring dashboards.",
        "questions": {
            "a": "Can the organization demonstrate: a system security plan is developed? Provide supporting evidence across all relevant components.",
            "b": "Can the organization demonstrate: the system boundary is described and documented in the system security plan? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: the system environment of operation is described and documented in the system security plan? Provide supporting evidence across all relevant components.",
            "d": "Has the organization identified the security requirements approved by the designated authority as non-applicable, and documented this determination in the system security plan?",
            "e": "Can the organization demonstrate: the method of implementation for each security requirement is described and documented in the system security plan? Provide supporting evidence across all relevant components.",
            "f": "Can the organization demonstrate: the relationship with or connection to other systems is described? Provide supporting evidence across all relevant components.",
            "g": "Has the organization defined the frequency to update the system security plan? Where is this documented?",
            "h": "Is the system security plan kept updated across servers, network devices, and workstations? What process/cadence ensures this?"
        }
    },
    "SC.L2-3.13.1": {
        "applies_to": "Network devices (firewalls, routers) at external and internal boundaries",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Has the organization defined the external system boundary? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Has the organization defined key internal system boundaries? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "c": "Is communications at the external system boundary monitored? What tool or process performs this monitoring (e.g., SIEM, network monitoring, EDR) and where are results reviewed?",
            "d": "Is communications at key internal boundaries monitored? What tool or process performs this monitoring (e.g., SIEM, network monitoring, EDR) and where are results reviewed?",
            "e": "How is communications at the external system boundary controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "f": "How is communications at key internal boundaries controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "g": "How is communications protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system).",
            "h": "How is communications protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system)."
        }
    },
    "SC.L2-3.13.2": {
        "applies_to": "Network devices, servers, cloud services",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Has the organization identified architectural designs that promote effective information security within organizational systems across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization identified software development techniques that promote effective information security within organizational systems across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "c": "Has the organization identified systems engineering principles that promote effective information security within organizational systems across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "d": "Can the organization demonstrate: identified architectural designs that promote effective information security are employed? Provide supporting evidence across all relevant components.",
            "e": "Can the organization demonstrate: identified software development techniques that promote effective information security are employed? Provide supporting evidence across all relevant components.",
            "f": "Can the organization demonstrate: identified systems engineering principles that promote effective information security are employed? Provide supporting evidence across all relevant components."
        }
    },
    "SC.L2-3.13.3": {
        "applies_to": "Network devices, servers, cloud services",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Has the organization identified user functionality across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization identified management functionality across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "c": "Can the organization demonstrate: user functionality is separated from management functionality? Provide supporting evidence across all relevant components."
        }
    },
    "SC.L2-3.13.4": {
        "applies_to": "Network devices, servers, cloud services",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Is unauthorized and unintended information transfer via shared system resources actively prevented/prohibited, and by what technical control (firewall rule, endpoint policy, etc.)?"
        }
    },
    "SC.L2-3.13.5": {
        "applies_to": "Network devices (DMZ / public-facing systems, network segmentation)",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Has the organization identified publicly accessible system components across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Can the organization demonstrate: subnetworks for publicly accessible system components are physically or logically separated from internal networks? Provide supporting evidence across all relevant components."
        }
    },
    "SC.L2-3.13.6": {
        "applies_to": "Network devices (firewalls, ACLs)",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Can the organization demonstrate: network communications traffic is denied by default? Provide supporting evidence across all relevant components.",
            "b": "Can the organization demonstrate: network communications traffic is allowed by exception? Provide supporting evidence across all relevant components."
        }
    },
    "SC.L2-3.13.7": {
        "applies_to": "Remote access infrastructure (VPN split-tunneling configuration)",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Are remote devices actually prevented from simultaneously establishing non-remote connections with the system and communicating via another connection to external network resources (split tunneling), except when securely provisioned? What technical control (firewall rule, endpoint policy, etc.) enforces this?"
        }
    },
    "SC.L2-3.13.8": {
        "applies_to": "Network devices, servers, cloud services",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Has the organization identified cryptographic mechanisms intended to prevent unauthorized disclosure of CUI across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Has the organization identified alternative physical safeguards intended to prevent unauthorized disclosure of CUI, for use where cryptographic mechanisms are not employed? Provide the current documented inventory.",
            "c": "Is identified cryptographic mechanisms or alternative physical safeguards actually implemented in practice, and how is this enforced across servers, network devices, workstations, and any third-party connections?"
        }
    },
    "SC.L2-3.13.9": {
        "applies_to": "Network devices, servers, cloud services",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Has the organization defined a period of inactivity to terminate network connections associated with communications sessions? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: network connections associated with communications sessions are terminated at the end of the sessions or after the defined period of inactivity? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: network connections associated with communications sessions are terminated after the defined period of inactivity? Provide supporting evidence across all relevant components."
        }
    },
    "SC.L2-3.13.10": {
        "applies_to": "Network devices, servers, cloud services",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Has cryptographic keys been established, and how is it maintained across servers, network devices, and the broader environment?",
            "b": "Can the organization demonstrate: cryptographic keys are managed whenever cryptography is employed? Provide supporting evidence across all relevant components."
        }
    },
    "SC.L2-3.13.11": {
        "applies_to": "Network devices, servers, cloud services",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Can the organization demonstrate: fIPS-validated cryptography is employed to protect the confidentiality of CUI? Provide supporting evidence across all relevant components."
        }
    },
    "SC.L2-3.13.12": {
        "applies_to": "Servers, network devices, collaboration platforms (video/voice conferencing)",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "Has the organization identified collaborative computing devices across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Can the organization demonstrate: collaborative computing devices provide indication to users of devices in use? Provide supporting evidence across all relevant components.",
            "c": "Is remote activation of collaborative computing devices actively prevented/prohibited, and by what technical control (firewall rule, endpoint policy, etc.)?"
        }
    },
    "SC.L2-3.13.13": {
        "applies_to": "Network devices, servers, cloud services",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "How is use of mobile code controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "b": "How is use of mobile code monitored, and by what mechanism or process — including for any third-party or vendor-managed components?"
        }
    },
    "SC.L2-3.13.14": {
        "applies_to": "Network devices, VoIP infrastructure",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "How is use of Voice over Internet Protocol (VoIP) technologies controlled, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "b": "How is use of Voice over Internet Protocol (VoIP) technologies monitored, and by what mechanism or process — including for any third-party or vendor-managed components?"
        }
    },
    "SC.L2-3.13.15": {
        "applies_to": "Network devices, servers, cloud services",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "How is the authenticity of communications sessions protected, and by what mechanism or process — including for any third-party or vendor-managed components?"
        }
    },
    "SC.L2-3.13.16": {
        "applies_to": "Network devices, servers, cloud services",
        "evidence_source": "Evidence: firewall/router/network device configuration (boundary protection, ACLs, segmentation), VPN and remote access configuration, TLS/encryption configuration for data in transit, server/database encryption at rest, third-party connection agreements (VPNs, APIs, extranets), DMZ/public-facing system separation.",
        "questions": {
            "a": "How is the confidentiality of CUI at rest protected? Identify the specific control on each relevant component (server, network device, workstation, or third-party system)."
        }
    },
    "SI.L2-3.14.1": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: endpoint/server antivirus-EDR configuration and update status, network device firmware/patch status, vulnerability/patch management program, IDS/IPS or network monitoring configuration, alerting integration (SIEM/SOC).",
        "questions": {
            "a": "Has the organization specified the time within which to identify system flaws? Where is this documented?",
            "b": "Has the organization identified system flaws across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "c": "Has the organization specified the time within which to report system flaws? Where is this documented?",
            "d": "Can the organization demonstrate: system flaws are reported within the specified time frame? Provide supporting evidence across all relevant components.",
            "e": "Has the organization specified the time within which to correct system flaws? Where is this documented?",
            "f": "Can the organization demonstrate: system flaws are corrected within the specified time frame? Provide supporting evidence across all relevant components."
        }
    },
    "SI.L2-3.14.2": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: endpoint/server antivirus-EDR configuration and update status, network device firmware/patch status, vulnerability/patch management program, IDS/IPS or network monitoring configuration, alerting integration (SIEM/SOC).",
        "questions": {
            "a": "Has the organization identified designated locations for malicious code protection across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "Can the organization demonstrate: protection from malicious code at designated locations is provided? Provide supporting evidence across all relevant components."
        }
    },
    "SI.L2-3.14.3": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: endpoint/server antivirus-EDR configuration and update status, network device firmware/patch status, vulnerability/patch management program, IDS/IPS or network monitoring configuration, alerting integration (SIEM/SOC).",
        "questions": {
            "a": "Has the organization identified response actions to system security alerts and advisories across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory.",
            "b": "How are system security alerts and advisories monitored, and by what mechanism or process — including for any third-party or vendor-managed components?",
            "c": "Can the organization demonstrate: identified response actions are taken in response to security alerts and advisories? Provide supporting evidence across all relevant components."
        }
    },
    "SI.L2-3.14.4": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: endpoint/server antivirus-EDR configuration and update status, network device firmware/patch status, vulnerability/patch management program, IDS/IPS or network monitoring configuration, alerting integration (SIEM/SOC).",
        "questions": {
            "a": "Is malicious code protection mechanisms kept updated across servers, network devices, and workstations? What process/cadence ensures this?"
        }
    },
    "SI.L2-3.14.5": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: endpoint/server antivirus-EDR configuration and update status, network device firmware/patch status, vulnerability/patch management program, IDS/IPS or network monitoring configuration, alerting integration (SIEM/SOC).",
        "questions": {
            "a": "Has the organization defined the frequency for malicious code scans? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Can the organization demonstrate: periodic scans of the system are performed with the defined frequency? Provide supporting evidence across all relevant components.",
            "c": "Can the organization demonstrate: real-time scans of files from external sources as files are downloaded, opened, or executed are performed? Provide supporting evidence across all relevant components."
        }
    },
    "SI.L2-3.14.6": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: endpoint/server antivirus-EDR configuration and update status, network device firmware/patch status, vulnerability/patch management program, IDS/IPS or network monitoring configuration, alerting integration (SIEM/SOC).",
        "questions": {
            "a": "Is the system monitored? What tool or process performs this monitoring (e.g., SIEM, network monitoring, EDR) and where are results reviewed?",
            "b": "Is inbound communications traffic monitored? What tool or process performs this monitoring (e.g., SIEM, network monitoring, EDR) and where are results reviewed?",
            "c": "Is outbound communications traffic monitored? What tool or process performs this monitoring (e.g., SIEM, network monitoring, EDR) and where are results reviewed?"
        }
    },
    "SI.L2-3.14.7": {
        "applies_to": "Servers, network devices, workstations",
        "evidence_source": "Evidence: endpoint/server antivirus-EDR configuration and update status, network device firmware/patch status, vulnerability/patch management program, IDS/IPS or network monitoring configuration, alerting integration (SIEM/SOC).",
        "questions": {
            "a": "Has the organization defined authorized use of the system? Where is this documented, and does it cover servers, network devices, and third-party connections in addition to workstations and cloud services?",
            "b": "Has the organization identified unauthorized use of the system across its full environment (servers, network devices, workstations, and third-party/external connections)? Provide the current documented inventory."
        }
    }
}

GAP_QUESTIONS_HAND_AUTHORED = [
    ('PS.L2-3.9.2', 'a'),
    ('PS.L2-3.9.2', 'b'),
    ('PE.L2-3.10.2', 'c'),
    ('PE.L2-3.10.2', 'd'),
    ('RA.L2-3.11.2', 'e'),
    ('CA.L2-3.12.2', 'c'),
    ('SC.L2-3.13.2', 'e'),
    ('SC.L2-3.13.2', 'f'),
    ('SC.L2-3.13.9', 'c'),
    ('SC.L2-3.13.13', 'b'),
    ('SC.L2-3.13.14', 'b'),
    ('SI.L2-3.14.2', 'a'),
    ('SI.L2-3.14.2', 'b'),
    ('SI.L2-3.14.3', 'b'),
    ('SI.L2-3.14.5', 'a'),
    ('MP.L2-3.8.2', 'a'),
    ('PS.L2-3.9.2', 'c'),
    ('PE.L2-3.10.2', 'a'),
    ('CA.L2-3.12.2', 'a'),
    ('CA.L2-3.12.4', 'd'),
    ('CA.L2-3.12.4', 'e'),
    ('CA.L2-3.12.4', 'g'),
    ('SC.L2-3.13.2', 'd'),
    ('SC.L2-3.13.7', 'a'),
    ('SC.L2-3.13.8', 'a'),
    ('SC.L2-3.13.8', 'b'),
    ('SC.L2-3.13.15', 'a'),
]
