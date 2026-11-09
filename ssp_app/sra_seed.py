"""Security & Risk Analysis reference content.

Sourced from the client-supplied Customer_Security_Risk_Analysis.xlsx
workbook: 222 assessment items across 3 control-set worksheets
(User Administrative Controls / I.T. Administrative Controls /
Training, HR, IRP & FSO), grounded in NIST SP 800-53 control families.
This is a distinct framework from the CMMC L2 / NIST SP 800-171
tooling elsewhere in this app (Controls, Gap Analysis, POA&M) — it is
not wired to CMMC_CONTROLS, and intentionally so, since the source
workbook assesses a different (broader, 800-53-based) control set.

Each item carries the pre-filled assessor guidance from the workbook's
'Examine' column where present; 'control_text' / 'objective_text' are
the control's own statement text, broken out the same way the source
workbook breaks a multi-part control (a/b/c..., 1/2/3...) across rows.
Item ids ('S<sheet#>-R<row#>') are derived from the source workbook's
own row positions so they stay stable across regenerations.

SRA_FAMILIES maps each NIST 800-53 family code to its canonical name
(polished from the workbook's abbreviated versions, e.g. 'Config Man'
-> 'Configuration Management'). SRA_SHEETS / SRA_SHEET_LABELS preserve
the workbook's 3 control-set groupings for the higher-level rollup the
source workbook's own Reports tab already showed.
"""

SRA_FAMILIES = {
    "AC": "Access Control",
    "AT": "Awareness and Training",
    "AU": "Audit and Accountability",
    "CM": "Configuration Management",
    "CP": "Contingency Planning",
    "IA": "Identification and Authentication",
    "IR": "Incident Response",
    "MP": "Media Protection",
    "PE": "Physical and Environmental Protection",
    "PL": "Planning",
    "PM": "Program Management",
    "PS": "Personnel Security",
    "RA": "Risk Assessment",
    "SA": "System and Services Acquisition",
    "SC": "System and Communications Protection",
    "SI": "System and Information Integrity"
}

SRA_SHEETS = [
    "1. User Admin. Controls",
    "2. I.T. Admin. Controls",
    "3. Training, HR, IRP, FSO"
]

SRA_SHEET_LABELS = {
    "1. User Admin. Controls": "User Administrative Controls",
    "2. I.T. Admin. Controls": "I.T. Administrative Controls",
    "3. Training, HR, IRP, FSO": "Training, HR, IRP & FSO"
}

SRA_ITEMS = [
    {
        "id": "S1-R2",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-6",
        "ref_name": "Least Privilege",
        "control_text": "Employ the principle of least privilege, allowing only authorized accesses for users (or processes acting on behalf of users) that are necessary to accomplish assigned organizational tasks.",
        "objective_text": "",
        "examine": "Is the Principle of Least Privileged maintained?"
    },
    {
        "id": "S1-R3",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-8",
        "ref_name": "System Use Notification",
        "control_text": "a.  Display [an organization-defined system use notification message or banner] to \nusers before granting access to the system that provides privacy and security notices \nconsistent with applicable laws, executive orders, directives, regulations, policies, standards, \nand guidelines and state that:",
        "objective_text": "1. Users are accessing a [Customer] managed system;",
        "examine": "Is there a Notice & Consent Banner or other type of notification utilized?"
    },
    {
        "id": "S1-R4",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-8",
        "ref_name": "System Use Notification",
        "control_text": "",
        "objective_text": "2. System usage may be monitored, recorded, and subject to audit;",
        "examine": "Does the banner (if any) contain similar verbiage?"
    },
    {
        "id": "S1-R5",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-8",
        "ref_name": "System Use Notification",
        "control_text": "",
        "objective_text": "3. Unauthorized use of the system is prohibited and subject to criminal and civil penalties; \nand",
        "examine": "Does the banner (if any) contain similar verbiage?"
    },
    {
        "id": "S1-R6",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-8",
        "ref_name": "System Use Notification",
        "control_text": "",
        "objective_text": "4. Use of the system indicates consent to monitoring and recording;",
        "examine": "Does the banner (if any) contain similar verbiage?"
    },
    {
        "id": "S1-R7",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-8",
        "ref_name": "System Use Notification",
        "control_text": "b.  Retain the notification message or banner on the screen until users acknowledge the usage \nconditions and take explicit actions to log on to or further access the system; and",
        "objective_text": "",
        "examine": "Does the banner (if any) implement a \"click-through\" or a \"click-on\" type of functionality to progress?"
    },
    {
        "id": "S1-R8",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-17",
        "ref_name": "Remote Access",
        "control_text": "a. Establish and document usage restrictions, configuration/connection requirements, and implementation guidance for each type of remote access allowed; and",
        "objective_text": "",
        "examine": "Does a Remote Access program exist?"
    },
    {
        "id": "S1-R9",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-17",
        "ref_name": "Remote Access",
        "control_text": "b. Authorize each type of remote access to the system prior to allowing such connections.",
        "objective_text": "",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R10",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-19",
        "ref_name": "Access Control For Mobile Devices",
        "control_text": "a. Establish configuration requirements, connection requirements, and implementation guidance for organization-controlled mobile devices, to include when such devices are outside of controlled areas; and",
        "objective_text": "",
        "examine": "Is a MAM or MDM type of program maintained?"
    },
    {
        "id": "S1-R11",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-19",
        "ref_name": "Access Control For Mobile Devices",
        "control_text": "b. Authorize the connection of mobile devices to organizational systems.",
        "objective_text": "",
        "examine": "Does MAM / MDM policy exist?"
    },
    {
        "id": "S1-R12",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-20",
        "ref_name": "Use of External Systems",
        "control_text": "a. [Selection (one or more): Establish [Assignment: organization-defined terms and conditions]; Identify [Assignment: organization-defined controls asserted to be implemented on external systems]], consistent with the trust relationships established with other organizations owning, operating, and/or maintaining external systems, allowing authorized individuals to:",
        "objective_text": "1. Access the system from external systems; and",
        "examine": "Are connection to external systems authorized and documented? I.e., personally owned devices, public kiosks, federal / non-federal information systems."
    },
    {
        "id": "S1-R13",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-20",
        "ref_name": "Use of External Systems",
        "control_text": "",
        "objective_text": "2. Process, store, or transmit organization-controlled information using external systems; or",
        "examine": "Is  authorized use maintained?"
    },
    {
        "id": "S1-R14",
        "sheet": "1. User Admin. Controls",
        "family": "AC",
        "ref": "AC-20",
        "ref_name": "Use of External Systems",
        "control_text": "b. Prohibit the use of [Assignment: organizationally-defined types of external systems].",
        "objective_text": "",
        "examine": "Are prohibitions outlined in policy?"
    },
    {
        "id": "S1-R15",
        "sheet": "1. User Admin. Controls",
        "family": "MP",
        "ref": "MP-7",
        "ref_name": "Media Use",
        "control_text": "a. [Selection: Restrict; Prohibit] the use of [Assignment: organization-defined types of system media] on [Assignment: organization-defined systems or system components] using [Assignment: organization-defined controls]; and",
        "objective_text": "",
        "examine": "Is a Portable Media type policy maintained?"
    },
    {
        "id": "S1-R16",
        "sheet": "1. User Admin. Controls",
        "family": "MP",
        "ref": "MP-7",
        "ref_name": "Media Use",
        "control_text": "b. Prohibit the use of portable storage devices in organizational systems when such devices have no identifiable owner.",
        "objective_text": "",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R17",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "a. Develop security and privacy plans for the system that:",
        "objective_text": "1. Are consistent with the organization’s enterprise architecture;",
        "examine": "Does a SSP type document exist?"
    },
    {
        "id": "S1-R18",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "2. Explicitly define the constituent system components;",
        "examine": "List of devices, processes / services, users and account associations (i.e., ACLs, PIM, RBAC)."
    },
    {
        "id": "S1-R19",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "3. Describe the operational context of the system in terms of mission and business processes;",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R20",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "4. Identify the individuals that fulfill system roles and responsibilities;",
        "examine": "List of devices, processes / services, users and account associations (i.e., ACLs, PIM, RBAC)."
    },
    {
        "id": "S1-R21",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "5. Identify the information types processed, stored, and transmitted by the system;",
        "examine": "All data types identified?"
    },
    {
        "id": "S1-R22",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "6. Provide the security categorization of the system, including supporting rationale;",
        "examine": "NIST SP 800-60 defines a four-step process for categorizing information and information systems as:\n\n(i) identify information types\n(ii) select provisional impact levels for the information types\n(iii) review provisional impact levels and adjust/finalize information impact levels for the information types\n(iv) assign a system security category and overall impact level."
    },
    {
        "id": "S1-R23",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "7. Describe any specific threats to the system that are of concern to the organization;",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R24",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "8. Provide the results of a privacy risk assessment for systems processing personally identifiable information;",
        "examine": "Was a Privacy Risk Assessment accomplished against the areas processing / storing CUI?"
    },
    {
        "id": "S1-R25",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "9. Describe the operational environment for the system and any dependencies on or connections to other systems or system components;",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R26",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "10. Provide an overview of the security and privacy requirements for the system;",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R27",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "11. Identify any relevant control baselines or overlays, if applicable;",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R28",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "12. Describe the controls in place or planned for meeting the security and privacy requirements, including a rationale for any tailoring decisions;",
        "examine": ""
    },
    {
        "id": "S1-R29",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "13. Include risk determinations for security and privacy architecture and design decisions;",
        "examine": ""
    },
    {
        "id": "S1-R30",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "14. Include security- and privacy-related activities affecting the system that require planning and coordination with [Assignment: organization-defined individuals or groups]; and",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R31",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "",
        "objective_text": "15. Are reviewed and approved by the authorizing official or designated representative prior to plan implementation.",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R32",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "b. Distribute copies of the plans and communicate subsequent changes to the plans to [Assignment: organization-defined personnel or roles];",
        "objective_text": "",
        "examine": "Does Executive Management review and approve (concur) with the SSP?"
    },
    {
        "id": "S1-R33",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "c. Review the plans [Assignment: organization-defined frequency];",
        "objective_text": "",
        "examine": "Is the SSP communicated and sponsored from a top-down management approach?"
    },
    {
        "id": "S1-R34",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "d. Update the plans to address changes to the system and environment of operation or problems identified during plan implementation or control assessments; and",
        "objective_text": "",
        "examine": "What is the defined frequency to review and approve the SSP?"
    },
    {
        "id": "S1-R35",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-2",
        "ref_name": "System Security and Privacy Plans",
        "control_text": "e. Protect the plans from unauthorized disclosure and modification.",
        "objective_text": "",
        "examine": "How are changes reviewed, approved, and incorporated into the SSP?"
    },
    {
        "id": "S1-R36",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-4",
        "ref_name": "Rules of Behavior",
        "control_text": "a. Establish and provide to individuals requiring access to the system, the rules that describe their responsibilities and expected behavior for information and system usage, security, and \nprivacy;",
        "objective_text": "",
        "examine": "Does an Acceptable Use type of policy exist for end-users, managers / supervisors, and I.T. / security managers?"
    },
    {
        "id": "S1-R37",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-4",
        "ref_name": "Rules of Behavior",
        "control_text": "b. Receive a documented acknowledgment from such individuals, indicating that they have \nread, understand, and agree to abide by the rules of behavior, before authorizing access to \ninformation and the system;",
        "objective_text": "",
        "examine": "How is concurrence to policy tracked / managed?"
    },
    {
        "id": "S1-R38",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-4",
        "ref_name": "Rules of Behavior",
        "control_text": "c. Review and update the rules of behavior [Assignment: organization-defined frequency]; and",
        "objective_text": "",
        "examine": "Is there a review process or frequency established to review and update this process?"
    },
    {
        "id": "S1-R39",
        "sheet": "1. User Admin. Controls",
        "family": "PL",
        "ref": "PL-4",
        "ref_name": "Rules of Behavior",
        "control_text": "d.  Require individuals who have acknowledged a previous version of the rules of behavior to \nread and re-acknowledge [Selection (one or more): [Assignment: organization-defined \nfrequency]; when the rules are revised or updated].",
        "objective_text": "1. When are the rules revised or updated?",
        "examine": "How is concurrence to policy tracked / managed?"
    },
    {
        "id": "S1-R40",
        "sheet": "1. User Admin. Controls",
        "family": "PS",
        "ref": "PS-6",
        "ref_name": "Access Agreements",
        "control_text": "a. Develop and document access agreements for organizational systems;",
        "objective_text": "",
        "examine": "Are System User Agreements or similar documentation maintained?"
    },
    {
        "id": "S1-R41",
        "sheet": "1. User Admin. Controls",
        "family": "PS",
        "ref": "PS-6",
        "ref_name": "Access Agreements",
        "control_text": "b. Review and update the access agreements [Assignment: organization-defined frequency]; and",
        "objective_text": "",
        "examine": "Is there a review process or frequence established to review and update this process?"
    },
    {
        "id": "S1-R42",
        "sheet": "1. User Admin. Controls",
        "family": "PS",
        "ref": "PS-6",
        "ref_name": "Access Agreements",
        "control_text": "c. Verify that individuals requiring access to organizational information and systems:",
        "objective_text": "1. Sign appropriate access agreements prior to being granted access; and",
        "examine": "Are System User Agreements or similar documentation maintained?"
    },
    {
        "id": "S1-R43",
        "sheet": "1. User Admin. Controls",
        "family": "PS",
        "ref": "PS-6",
        "ref_name": "Access Agreements",
        "control_text": "",
        "objective_text": "2. Re-sign access agreements to maintain access to organizational systems when access \nagreements have been updated or [Assignment: organization-defined frequency].",
        "examine": "Is there a review process or frequence established to review and update this process?"
    },
    {
        "id": "S1-R44",
        "sheet": "1. User Admin. Controls",
        "family": "RA",
        "ref": "RA-8",
        "ref_name": "Privacy Impact Assessments",
        "control_text": "Conduct privacy impact assessments for systems, programs, or other activities before:",
        "objective_text": "",
        "examine": "Is a Privacy Impact type assessment maintained?"
    },
    {
        "id": "S1-R45",
        "sheet": "1. User Admin. Controls",
        "family": "RA",
        "ref": "RA-8",
        "ref_name": "Privacy Impact Assessments",
        "control_text": "a. Developing or procuring information technology that processes personally identifiable information; and",
        "objective_text": "",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R46",
        "sheet": "1. User Admin. Controls",
        "family": "RA",
        "ref": "RA-8",
        "ref_name": "Privacy Impact Assessments",
        "control_text": "b. Initiating a new collection of personally identifiable information that:",
        "objective_text": "1. Will be processed using information technology; and",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R47",
        "sheet": "1. User Admin. Controls",
        "family": "RA",
        "ref": "RA-8",
        "ref_name": "Privacy Impact Assessments",
        "control_text": "",
        "objective_text": "2. Includes personally identifiable information permitting the physical or virtual (online) contacting of a specific individual, if identical questions have been posed to, or identical reporting requirements imposed on, ten or more individuals, other than agencies, instrumentalities, or employees of the federal government.",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S1-R48",
        "sheet": "1. User Admin. Controls",
        "family": "SC",
        "ref": "SC-28",
        "ref_name": "Protection Of Information At Rest",
        "control_text": "Protect the [Selection (one or more): confidentiality; integrity] of the following information at rest: [Assignment: organization-defined information at rest].",
        "objective_text": "",
        "examine": "Is Data-At-Rest (DAR) implemented?"
    },
    {
        "id": "S1-R49",
        "sheet": "1. User Admin. Controls",
        "family": "SI",
        "ref": "SI-3",
        "ref_name": "Malicious Code Protection",
        "control_text": "a. Implement [Selection (one or more): signature based; non-signature based] malicious code protection mechanisms at system entry and exit points to detect and eradicate malicious code;",
        "objective_text": "",
        "examine": "Is a Malware protection program maintained?"
    },
    {
        "id": "S1-R50",
        "sheet": "1. User Admin. Controls",
        "family": "SI",
        "ref": "SI-3",
        "ref_name": "Malicious Code Protection",
        "control_text": "b. Automatically update malicious code protection mechanisms as new releases are available in accordance with organizational configuration management policy and procedures;",
        "objective_text": "",
        "examine": "Is Malware protection tools / mechanisms updated upon new release / updates?"
    },
    {
        "id": "S1-R51",
        "sheet": "1. User Admin. Controls",
        "family": "SI",
        "ref": "SI-3",
        "ref_name": "Malicious Code Protection",
        "control_text": "c. Configure malicious code protection mechanisms to:",
        "objective_text": "1. Perform periodic scans of the system [Assignment: organization-defined frequency] and real-time scans of files from external sources at [Selection (one or more): endpoint; network entry and exit points] as the files are downloaded, opened, or executed in accordance with organizational policy; and",
        "examine": "Determine the configuration."
    },
    {
        "id": "S1-R52",
        "sheet": "1. User Admin. Controls",
        "family": "SI",
        "ref": "SI-3",
        "ref_name": "Malicious Code Protection",
        "control_text": "",
        "objective_text": "2. [Selection (one or more): block malicious code; quarantine malicious code; take [Assignment: organization-defined action]]; and send alert to [Assignment: organization-defined personnel or roles] in response to malicious code detection; and",
        "examine": "Determine the configuration."
    },
    {
        "id": "S1-R53",
        "sheet": "1. User Admin. Controls",
        "family": "SI",
        "ref": "SI-3",
        "ref_name": "Malicious Code Protection",
        "control_text": "d. Address the receipt of false positives during malicious code detection and eradication and the resulting potential impact on the availability of the system.",
        "objective_text": "",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S2-R2",
        "sheet": "2. I.T. Admin. Controls",
        "family": "AC",
        "ref": "AC-4",
        "ref_name": "Information Flow Enforcement",
        "control_text": "Enforce approved authorizations for controlling the flow of information within the system and between connected systems based on [Assignment: organization-defined information flow control policies].",
        "objective_text": "",
        "examine": "Are there data flow restrictions or controls utilized?"
    },
    {
        "id": "S2-R3",
        "sheet": "2. I.T. Admin. Controls",
        "family": "AU",
        "ref": "AU-6",
        "ref_name": "Audit Record Review, Analysis. And Reporting",
        "control_text": "a. Review and analyze system audit records [Assignment: organization-defined frequency] for \nindications of [Assignment: organization-defined inappropriate or unusual activity] and the \npotential impact of the inappropriate or unusual activity;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R4",
        "sheet": "2. I.T. Admin. Controls",
        "family": "AU",
        "ref": "AU-6",
        "ref_name": "Audit Record Review, Analysis. And Reporting",
        "control_text": "b. Report findings to [Assignment: organization-defined personnel or roles]; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R5",
        "sheet": "2. I.T. Admin. Controls",
        "family": "AU",
        "ref": "AU-6",
        "ref_name": "Audit Record Review, Analysis. And Reporting",
        "control_text": "c. Adjust the level of audit record review, analysis, and reporting within the system when there \nis a change in risk based on law enforcement information, intelligence information, or other \ncredible sources of information.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R6",
        "sheet": "2. I.T. Admin. Controls",
        "family": "AU",
        "ref": "AU-12",
        "ref_name": "Audit Record Generation",
        "control_text": "a. Provide audit record generation capability for the event types the system is capable of \nauditing as defined in AU-2a on [Assignment: organization-defined system components];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R7",
        "sheet": "2. I.T. Admin. Controls",
        "family": "AU",
        "ref": "AU-12",
        "ref_name": "Audit Record Generation",
        "control_text": "b. Allow [Assignment: organization-defined personnel or roles] to select the event types that \nare to be logged by specific components of the system; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R8",
        "sheet": "2. I.T. Admin. Controls",
        "family": "AU",
        "ref": "AU-12",
        "ref_name": "Audit Record Generation",
        "control_text": "c. Generate audit records for the event types defined in AU-2c that include the audit record\ncontent defined in AU-3",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R9",
        "sheet": "2. I.T. Admin. Controls",
        "family": "CM",
        "ref": "CM-2",
        "ref_name": "Baseline Configuration",
        "control_text": "a. Develop, document, and maintain under configuration control, a current baseline configuration of the system; and",
        "objective_text": "",
        "examine": "Does a SYSTEM CONFIGURATION BASELINE exist? Does the baseline include documented controls, authorizations, and restrictions?"
    },
    {
        "id": "S2-R10",
        "sheet": "2. I.T. Admin. Controls",
        "family": "CM",
        "ref": "CM-2",
        "ref_name": "Baseline Configuration",
        "control_text": "b. Review and update the baseline configuration of the system:",
        "objective_text": "1. [Assignment: organization-defined frequency];",
        "examine": "Are restrictions or controls documented and authorized?"
    },
    {
        "id": "S2-R11",
        "sheet": "2. I.T. Admin. Controls",
        "family": "CM",
        "ref": "CM-2",
        "ref_name": "Baseline Configuration",
        "control_text": "",
        "objective_text": "2. When required due to [Assignment: organization-defined circumstances]; and",
        "examine": "Are restrictions or controls documented and authorized?"
    },
    {
        "id": "S2-R12",
        "sheet": "2. I.T. Admin. Controls",
        "family": "CM",
        "ref": "CM-2",
        "ref_name": "Baseline Configuration",
        "control_text": "",
        "objective_text": "3. When system components are installed or upgraded.",
        "examine": "Are restrictions or controls documented and authorized?"
    },
    {
        "id": "S2-R13",
        "sheet": "2. I.T. Admin. Controls",
        "family": "CM",
        "ref": "CM-6",
        "ref_name": "Configuration Settings",
        "control_text": "a. Establish and document configuration settings for components employed within the system \nthat reflect the most restrictive mode consistent with operational requirements using \n[Assignment: organization-defined common secure configurations];",
        "objective_text": "",
        "examine": "Does a SYSTEM-COMPONENTS CONFIGURATION BASELINE exist?"
    },
    {
        "id": "S2-R14",
        "sheet": "2. I.T. Admin. Controls",
        "family": "CM",
        "ref": "CM-6",
        "ref_name": "Configuration Settings",
        "control_text": "b. Implement the configuration settings;",
        "objective_text": "",
        "examine": "Does the baseline address device configuration as abiding to the Principle of Least Functionality (e.g., CM-7)?"
    },
    {
        "id": "S2-R15",
        "sheet": "2. I.T. Admin. Controls",
        "family": "CM",
        "ref": "CM-6",
        "ref_name": "Configuration Settings",
        "control_text": "c. Identify, document, and approve any deviations from established configuration settings for \n[Assignment: organization-defined system components] based on [Assignment: organization defined operational requirements]; and",
        "objective_text": "",
        "examine": "Is a Change Review type program maintained?"
    },
    {
        "id": "S2-R16",
        "sheet": "2. I.T. Admin. Controls",
        "family": "CM",
        "ref": "CM-6",
        "ref_name": "Configuration Settings",
        "control_text": "d. Monitor and control changes to the configuration settings in accordance with organizational \npolicies and procedures.",
        "objective_text": "",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S2-R17",
        "sheet": "2. I.T. Admin. Controls",
        "family": "CM",
        "ref": "CM-7",
        "ref_name": "Least Functionality",
        "control_text": "a. Configure the system to provide only [Assignment: organization-defined mission essential capabilities]; and",
        "objective_text": "",
        "examine": "Does the baseline address device configuration as abiding to the Principle of Least Functionality?"
    },
    {
        "id": "S2-R18",
        "sheet": "2. I.T. Admin. Controls",
        "family": "CM",
        "ref": "CM-7",
        "ref_name": "Least Functionality",
        "control_text": "b. Prohibit or restrict the use of the following functions, ports, protocols, software, and/or services: [Assignment: organization-defined prohibited or restricted functions, system ports, protocols, software, and/or services].",
        "objective_text": "",
        "examine": "Does policy include the necessary verbiage?"
    },
    {
        "id": "S2-R19",
        "sheet": "2. I.T. Admin. Controls",
        "family": "IA",
        "ref": "IA-2",
        "ref_name": "Identification and Authentication",
        "control_text": "Control: Uniquely identify and authenticate organizational users and associate that unique identification with processes acting on behalf of those users.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R20",
        "sheet": "2. I.T. Admin. Controls",
        "family": "MP",
        "ref": "MP-4",
        "ref_name": "Media Storage",
        "control_text": "a. Physically control and securely store [Assignment: organization-defined types of digital and/or non-digital media] within [Assignment: organization-defined controlled areas]; and",
        "objective_text": "",
        "examine": "Is there a media storage program?"
    },
    {
        "id": "S2-R21",
        "sheet": "2. I.T. Admin. Controls",
        "family": "MP",
        "ref": "MP-4",
        "ref_name": "Media Storage",
        "control_text": "b. Protect system media types defined in MP-4a until the media are destroyed or sanitized using approved equipment, techniques, and procedures.",
        "objective_text": "",
        "examine": "Does the program include a sanitation type program?"
    },
    {
        "id": "S2-R22",
        "sheet": "2. I.T. Admin. Controls",
        "family": "MP",
        "ref": "MP-5",
        "ref_name": "Media Transport",
        "control_text": "a. Protect and control [Assignment: organization-defined types of system media] during transport outside of controlled areas using [Assignment: organization-defined controls];",
        "objective_text": "",
        "examine": "Is there a media handling program?"
    },
    {
        "id": "S2-R23",
        "sheet": "2. I.T. Admin. Controls",
        "family": "MP",
        "ref": "MP-5",
        "ref_name": "Media Transport",
        "control_text": "b. Maintain accountability for system media during transport outside of controlled areas;",
        "objective_text": "",
        "examine": "Does the program include accountability?"
    },
    {
        "id": "S2-R24",
        "sheet": "2. I.T. Admin. Controls",
        "family": "MP",
        "ref": "MP-5",
        "ref_name": "Media Transport",
        "control_text": "c. Document activities associated with the transport of system media; and",
        "objective_text": "",
        "examine": "Does similar verbiage exist in policy?"
    },
    {
        "id": "S2-R25",
        "sheet": "2. I.T. Admin. Controls",
        "family": "MP",
        "ref": "MP-5",
        "ref_name": "Media Transport",
        "control_text": "d. Restrict the activities associated with the transport of system media to authorized personnel.",
        "objective_text": "",
        "examine": "Does similar verbiage exist in policy?"
    },
    {
        "id": "S2-R26",
        "sheet": "2. I.T. Admin. Controls",
        "family": "MP",
        "ref": "MP-6",
        "ref_name": "Media Sanitization",
        "control_text": "a. Sanitize [Assignment: organization-defined system media] prior to disposal, release out of \norganizational control, or release for reuse using [Assignment: organization-defined \nsanitization techniques and procedures]; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R27",
        "sheet": "2. I.T. Admin. Controls",
        "family": "MP",
        "ref": "MP-6",
        "ref_name": "Media Sanitization",
        "control_text": "b. Employ sanitization mechanisms with the strength and integrity commensurate with the \nsecurity category or classification of the information.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R28",
        "sheet": "2. I.T. Admin. Controls",
        "family": "PE",
        "ref": "PE-9",
        "ref_name": "Power Equipment and Cabling",
        "control_text": "Protect power equipment and power cabling for the system from damage and destruction.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R29",
        "sheet": "2. I.T. Admin. Controls",
        "family": "PM",
        "ref": "PM-1",
        "ref_name": "Information Security Program Plan",
        "control_text": "a. Develop and disseminate an organization-wide information security program plan that:",
        "objective_text": "1. Provides an overview of the requirements for the security program and a description of \nthe security program management controls and common controls in place or planned \nfor meeting those requirements;",
        "examine": ""
    },
    {
        "id": "S2-R30",
        "sheet": "2. I.T. Admin. Controls",
        "family": "PM",
        "ref": "PM-1",
        "ref_name": "Information Security Program Plan",
        "control_text": "",
        "objective_text": "2. Includes the identification and assignment of roles, responsibilities, management \ncommitment, coordination among organizational entities, and compliance;",
        "examine": ""
    },
    {
        "id": "S2-R31",
        "sheet": "2. I.T. Admin. Controls",
        "family": "PM",
        "ref": "PM-1",
        "ref_name": "Information Security Program Plan",
        "control_text": "",
        "objective_text": "3. Reflects the coordination among organizational entities responsible for information \nsecurity; and",
        "examine": ""
    },
    {
        "id": "S2-R32",
        "sheet": "2. I.T. Admin. Controls",
        "family": "PM",
        "ref": "PM-1",
        "ref_name": "Information Security Program Plan",
        "control_text": "",
        "objective_text": "4. Is approved by a senior official with responsibility and accountability for the risk being \nincurred to organizational operations (including mission, functions, image, and \nreputation), organizational assets, individuals, other organizations, and the Nation;",
        "examine": ""
    },
    {
        "id": "S2-R33",
        "sheet": "2. I.T. Admin. Controls",
        "family": "PM",
        "ref": "PM-1",
        "ref_name": "Information Security Program Plan",
        "control_text": "b. Review and update the organization-wide information security program plan [Assignment: \norganization-defined frequency] and following [Assignment: organization-defined events];\nand",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R34",
        "sheet": "2. I.T. Admin. Controls",
        "family": "PM",
        "ref": "PM-1",
        "ref_name": "Information Security Program Plan",
        "control_text": "c. Protect the information security program plan from unauthorized disclosure and \nmodification",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R35",
        "sheet": "2. I.T. Admin. Controls",
        "family": "RA",
        "ref": "RA-5",
        "ref_name": "Vulnerability Monitoring and Scanning",
        "control_text": "a. Monitor and scan for vulnerabilities in the system and hosted applications [Assignment: \norganization-defined frequency and/or randomly in accordance with organization-defined \nprocess] and when new vulnerabilities potentially affecting the system are identified and \nreported;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R36",
        "sheet": "2. I.T. Admin. Controls",
        "family": "RA",
        "ref": "RA-5",
        "ref_name": "Vulnerability Monitoring and Scanning",
        "control_text": "b. Employ vulnerability monitoring tools and techniques that facilitate interoperability among \ntools and automate parts of the vulnerability management process by using standards for:",
        "objective_text": "1. Enumerating platforms, software flaws, and improper configurations;",
        "examine": ""
    },
    {
        "id": "S2-R37",
        "sheet": "2. I.T. Admin. Controls",
        "family": "RA",
        "ref": "RA-5",
        "ref_name": "Vulnerability Monitoring and Scanning",
        "control_text": "",
        "objective_text": "2. Formatting checklists and test procedures; and",
        "examine": ""
    },
    {
        "id": "S2-R38",
        "sheet": "2. I.T. Admin. Controls",
        "family": "RA",
        "ref": "RA-5",
        "ref_name": "Vulnerability Monitoring and Scanning",
        "control_text": "",
        "objective_text": "3. Measuring vulnerability impact;",
        "examine": ""
    },
    {
        "id": "S2-R39",
        "sheet": "2. I.T. Admin. Controls",
        "family": "RA",
        "ref": "RA-5",
        "ref_name": "Vulnerability Monitoring and Scanning",
        "control_text": "c. Analyze vulnerability scan reports and results from vulnerability monitoring;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R40",
        "sheet": "2. I.T. Admin. Controls",
        "family": "RA",
        "ref": "RA-5",
        "ref_name": "Vulnerability Monitoring and Scanning",
        "control_text": "d. Remediate legitimate vulnerabilities [Assignment: organization-defined response times] in accordance with an organizational assessment of risk;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R41",
        "sheet": "2. I.T. Admin. Controls",
        "family": "RA",
        "ref": "RA-5",
        "ref_name": "Vulnerability Monitoring and Scanning",
        "control_text": "e. Share information obtained from the vulnerability monitoring process and control assessments with [Assignment: organization-defined personnel or roles] to help eliminate similar vulnerabilities in other systems; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R42",
        "sheet": "2. I.T. Admin. Controls",
        "family": "RA",
        "ref": "RA-5",
        "ref_name": "Vulnerability Monitoring and Scanning",
        "control_text": "f. Employ vulnerability monitoring tools that include the capability to readily update the vulnerabilities to be scanned.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R43",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SC",
        "ref": "SC-7",
        "ref_name": "Boundary Protection",
        "control_text": "a. Monitor and control communications at the external managed interfaces to the system and at key internal managed interfaces within the system;",
        "objective_text": "",
        "examine": "Network security type program utilized?"
    },
    {
        "id": "S2-R44",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SC",
        "ref": "SC-7",
        "ref_name": "Boundary Protection",
        "control_text": "b. Implement subnetworks for publicly accessible system components that are [Selection: physically; logically] separated from internal organizational networks; and",
        "objective_text": "",
        "examine": "Segmentation occurring, network boundaries established, baseline separation occurring?"
    },
    {
        "id": "S2-R45",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SC",
        "ref": "SC-7",
        "ref_name": "Boundary Protection",
        "control_text": "c. Connect to external networks or systems only through managed interfaces consisting of boundary protection devices arranged in accordance with an organizational security and privacy architecture.",
        "objective_text": "",
        "examine": "Controlled and managed access points utilized or implemented?"
    },
    {
        "id": "S2-R46",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SC",
        "ref": "SC-13",
        "ref_name": "Cryptographic Protection",
        "control_text": "a. Determine the [Assignment: organization-defined cryptographic uses]; and",
        "objective_text": "",
        "examine": "List the types of encryption devices and the encryption module in use per device (if any)."
    },
    {
        "id": "S2-R47",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SC",
        "ref": "SC-13",
        "ref_name": "Cryptographic Protection",
        "control_text": "b. Implement the following types of cryptography required for each specified cryptographic use: [Assignment: organization-defined types of cryptography for each specified cryptographic use].",
        "objective_text": "",
        "examine": "List the types of encryption devices and the encryption module in use per device (if any)."
    },
    {
        "id": "S2-R48",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-4",
        "ref_name": "System Monitoring",
        "control_text": "a. Monitor the system to detect:",
        "objective_text": "1. Attacks and indicators of potential attacks in accordance with the following monitoring objectives: [Assignment: organization-defined monitoring objectives]; and",
        "examine": "Is an IDS type mechanism employed and configured?"
    },
    {
        "id": "S2-R49",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-4",
        "ref_name": "System Monitoring",
        "control_text": "",
        "objective_text": "2. Unauthorized local, network, and remote connections;",
        "examine": "Is an IDS type mechanism employed and configured?"
    },
    {
        "id": "S2-R50",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-4",
        "ref_name": "System Monitoring",
        "control_text": "b. Identify unauthorized use of the system through the following techniques and methods: [Assignment: organization-defined techniques and methods];",
        "objective_text": "",
        "examine": "Is an IDS type mechanism employed and configured?"
    },
    {
        "id": "S2-R51",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-4",
        "ref_name": "System Monitoring",
        "control_text": "c. Invoke internal monitoring capabilities or deploy monitoring devices:",
        "objective_text": "1. Strategically within the system to collect organization-determined essential information; and",
        "examine": "Is an IDS type mechanism employed and configured?"
    },
    {
        "id": "S2-R52",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-4",
        "ref_name": "System Monitoring",
        "control_text": "",
        "objective_text": "2. At ad hoc locations within the system to track specific types of transactions of interest to the organization;",
        "examine": "Is an IDS type mechanism employed and configured?"
    },
    {
        "id": "S2-R53",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-4",
        "ref_name": "System Monitoring",
        "control_text": "d. Analyze detected events and anomalies;",
        "objective_text": "",
        "examine": "Is there a mechanism to aggregate detected events and anomalies for analysis?"
    },
    {
        "id": "S2-R54",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-4",
        "ref_name": "System Monitoring",
        "control_text": "e. Adjust the level of system monitoring activity when there is a change in risk to organizational operations and assets, individuals, other organizations, or the Nation;",
        "objective_text": "",
        "examine": "Are restrictions or controls documented and authorized?"
    },
    {
        "id": "S2-R55",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-4",
        "ref_name": "System Monitoring",
        "control_text": "f. Obtain legal opinion regarding system monitoring activities; and",
        "objective_text": "",
        "examine": "Are restrictions or controls documented and authorized?"
    },
    {
        "id": "S2-R56",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-4",
        "ref_name": "System Monitoring",
        "control_text": "g. Provide [Assignment: organization-defined system monitoring information] to [Assignment: organization-defined personnel or roles] [Selection (one or more): as needed; [Assignment: organization-defined frequency]].",
        "objective_text": "",
        "examine": "Are restrictions or controls documented and authorized?"
    },
    {
        "id": "S2-R57",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-2",
        "ref_name": "Flaw Remediation",
        "control_text": "a. Identify, report, and correct system flaws;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R58",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-2",
        "ref_name": "Flaw Remediation",
        "control_text": "c. Install security-relevant software and firmware updates within [Assignment: organization defined time period] of the release of the updates",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S2-R59",
        "sheet": "2. I.T. Admin. Controls",
        "family": "SI",
        "ref": "SI-2",
        "ref_name": "Flaw Remediation",
        "control_text": "d. Incorporate flaw remediation into the organizational configuration management process.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R2",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "a. Define and document the types of accounts allowed and specifically prohibited for use within the system;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R3",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "b. Assign account managers;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R4",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "c. Require [Assignment: organization-defined prerequisites and criteria] for group and role membership;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R5",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "d. Specify:",
        "objective_text": "1. Authorized users of the system;",
        "examine": ""
    },
    {
        "id": "S3-R6",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "",
        "objective_text": "2. Group and role membership; and",
        "examine": ""
    },
    {
        "id": "S3-R7",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "",
        "objective_text": "3. Access authorizations (i.e., privileges) and [Assignment: organization-defined attributes (as required)] for each account;",
        "examine": ""
    },
    {
        "id": "S3-R8",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "e. Require approvals by [Assignment: organization-defined personnel or roles] for requests to create accounts;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R9",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "f. Create, enable, modify, disable, and remove accounts in accordance with [Assignment: organization-defined policy, procedures, prerequisites, and criteria];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R10",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "g. Monitor the use of accounts;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R11",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "h. Notify account managers and [Assignment: organization-defined personnel or roles] within:",
        "objective_text": "1. [Assignment: organization-defined time period] when accounts are no longer required;",
        "examine": ""
    },
    {
        "id": "S3-R12",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "",
        "objective_text": "2. [Assignment: organization-defined time period] when users are terminated or transferred; and",
        "examine": ""
    },
    {
        "id": "S3-R13",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "",
        "objective_text": "3. [Assignment: organization-defined time period] when system usage or need-to-know changes for an individual;",
        "examine": ""
    },
    {
        "id": "S3-R14",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "i. Authorize access to the system based on:",
        "objective_text": "1. A valid access authorization;",
        "examine": ""
    },
    {
        "id": "S3-R15",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "",
        "objective_text": "2. Intended system usage; and",
        "examine": ""
    },
    {
        "id": "S3-R16",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "",
        "objective_text": "3. [Assignment: organization-defined attributes (as required)];",
        "examine": ""
    },
    {
        "id": "S3-R17",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "j. Review accounts for compliance with account management requirements [Assignment: organization-defined frequency];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R18",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "k. Establish and implement a process for changing shared or group account authenticators (if deployed) when individuals are removed from the group; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R19",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-2",
        "ref_name": "Account Management",
        "control_text": "l. Align account management processes with personnel termination and transfer processes.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R20",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-5",
        "ref_name": "Separation of Duties",
        "control_text": "a. Identify and document [Assignment: organization-defined duties of individuals requiring \nseparation]; and",
        "objective_text": "",
        "examine": "Are roles and permissions defined and are personnel assigned to authorized individuals? E.g., Documented and approved."
    },
    {
        "id": "S3-R21",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AC",
        "ref": "AC-5",
        "ref_name": "Separation of Duties",
        "control_text": "b. Define system access authorizations to support separation of duties.",
        "objective_text": "",
        "examine": "Are roles and permissions defined and are personnel assigned to authorized individuals? E.g., Documented and approved."
    },
    {
        "id": "S3-R22",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AT",
        "ref": "AT-2",
        "ref_name": "Literacy Training And Awareness,",
        "control_text": "a. Provide security and privacy literacy training to system users (including managers, senior \nexecutives, and contractors)",
        "objective_text": "1. As part of initial training for new users and [Assignment: organization-defined \nfrequency] thereafter; and",
        "examine": "Describe the Security Awareness Training program. Does training include baseline security awareness and privacy awareness type training? Is this initial and continuous training or one-time only?"
    },
    {
        "id": "S3-R23",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AT",
        "ref": "AT-2",
        "ref_name": "Literacy Training And Awareness,",
        "control_text": "",
        "objective_text": "2. When required by system changes or following [Assignment: organization-defined events];",
        "examine": "When (if any) is training required / mandated?"
    },
    {
        "id": "S3-R24",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AT",
        "ref": "AT-2",
        "ref_name": "Literacy Training And Awareness,",
        "control_text": "b. Employ the following techniques to increase the security and privacy awareness of system \nusers [Assignment: organization-defined awareness techniques]",
        "objective_text": "",
        "examine": "Is there specific types of training unique to the organization?"
    },
    {
        "id": "S3-R25",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AT",
        "ref": "AT-2",
        "ref_name": "Literacy Training And Awareness,",
        "control_text": "c. Update literacy training and awareness content [Assignment: organization-defined frequency] and following [Assignment: organization-defined events]; and",
        "objective_text": "",
        "examine": "How often is training updated or revised? When was the last update or revision?"
    },
    {
        "id": "S3-R26",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AT",
        "ref": "AT-2",
        "ref_name": "Literacy Training And Awareness,",
        "control_text": "d. Incorporate lessons learned from internal or external security incidents or breaches into \nliteracy training and awareness techniques.",
        "objective_text": "",
        "examine": "Does training include industry best practices and lessons learned based on the current environmental threats?"
    },
    {
        "id": "S3-R27",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AT",
        "ref": "AT-3",
        "ref_name": "Role-Based Training",
        "control_text": "a. Provide role-based security and privacy training to personnel with the following roles and \nresponsibilities: [Assignment: organization-defined roles and responsibilities]:",
        "objective_text": "1. Before authorizing access to the system, information, or performing assigned duties, \nand [Assignment: organization-defined frequency] thereafter; and",
        "examine": "Does training incorporate RBAC, ACL, PIM, or other privileged user awareness type training?"
    },
    {
        "id": "S3-R28",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AT",
        "ref": "AT-3",
        "ref_name": "Role-Based Training",
        "control_text": "",
        "objective_text": "2. When required by system changes;",
        "examine": "When (if any) is training required / mandated?"
    },
    {
        "id": "S3-R29",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AT",
        "ref": "AT-3",
        "ref_name": "Role-Based Training",
        "control_text": "b. Update role-based training content [Assignment: organization-defined frequency] and \nfollowing [Assignment: organization-defined events]; and",
        "objective_text": "",
        "examine": "How often is training updated or revised? When was the last update or revision?"
    },
    {
        "id": "S3-R30",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "AT",
        "ref": "AT-3",
        "ref_name": "Role-Based Training",
        "control_text": "c. Incorporate lessons learned from internal or external security incidents or breaches into \nrole-based training.",
        "objective_text": "",
        "examine": "Does training include industry best practices and lessons learned based on the current environmental threats?"
    },
    {
        "id": "S3-R31",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "a. Develop a contingency plan for the system that:",
        "objective_text": "1. Identifies essential mission and business functions and associated contingency \nrequirements;",
        "examine": ""
    },
    {
        "id": "S3-R32",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "",
        "objective_text": "2. Provides recovery objectives, restoration priorities, and metrics;",
        "examine": ""
    },
    {
        "id": "S3-R33",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "",
        "objective_text": "3. Addresses contingency roles, responsibilities, assigned individuals with contact \ninformation;",
        "examine": ""
    },
    {
        "id": "S3-R34",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "",
        "objective_text": "4. Addresses maintaining essential mission and business functions despite a system \ndisruption, compromise, or failure;",
        "examine": ""
    },
    {
        "id": "S3-R35",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "",
        "objective_text": "5. Addresses eventual, full system restoration without deterioration of the controls \noriginally planned and implemented;",
        "examine": ""
    },
    {
        "id": "S3-R36",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "",
        "objective_text": "6. Addresses the sharing of contingency information; and",
        "examine": ""
    },
    {
        "id": "S3-R37",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "",
        "objective_text": "7. Is reviewed and approved by [Assignment: organization-defined personnel or roles];",
        "examine": ""
    },
    {
        "id": "S3-R38",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "b. Distribute copies of the contingency plan to [Assignment: organization-defined key \ncontingency personnel (identified by name and/or by role) and organizational elements];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R39",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "c. Coordinate contingency planning activities with incident handling activities;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R40",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "d. Review the contingency plan for the system [Assignment: organization-defined frequency]",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R41",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "e. Update the contingency plan to address changes to the organization, system, or \nenvironment of operation and problems encountered during contingency plan \nimplementation, execution, or testing;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R42",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "f. Communicate contingency plan changes to [Assignment: organization-defined key \ncontingency personnel (identified by name and/or by role) and organizational elements];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R43",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "g. Incorporate lessons learned from contingency plan testing, training, or actual contingency\nactivities into contingency testing and training; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R44",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-2",
        "ref_name": "Contingency Plan",
        "control_text": "h. Protect the contingency plan from unauthorized disclosure and modification.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R45",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-6",
        "ref_name": "Alternate Storage Site",
        "control_text": "a. Conduct backups of user-level information contained in [Assignment: organization-defined \nsystem components] [Assignment: organization-defined frequency consistent with recovery \ntime and recovery point objectives]",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R46",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-6",
        "ref_name": "Alternate Storage Site",
        "control_text": "b. Conduct backups of system-level information contained in the system [Assignment: \norganization-defined frequency consistent with recovery time and recovery point objectives]",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R47",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-6",
        "ref_name": "Alternate Storage Site",
        "control_text": "c. Conduct backups of system documentation, including security- and privacy-related \ndocumentation [Assignment: organization-defined frequency consistent with recovery time \nand recovery point objectives]; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R48",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "CP",
        "ref": "CP-6",
        "ref_name": "Alternate Storage Site",
        "control_text": "d. Protect the confidentiality, integrity, and availability of backup information.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R49",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IA",
        "ref": "IA-4",
        "ref_name": "Identifier Management",
        "control_text": "Manage system identifiers by:\na. Receiving authorization from [Assignment: organization-defined personnel or roles] to assign an individual, group, role, service, or device identifier;",
        "objective_text": "",
        "examine": "Are authorizations / approvals required for account creations?"
    },
    {
        "id": "S3-R50",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IA",
        "ref": "IA-4",
        "ref_name": "Identifier Management",
        "control_text": "b. Selecting an identifier that identifies an individual, group, role, service, or device;",
        "objective_text": "",
        "examine": "How are user accounts defined / controlled?"
    },
    {
        "id": "S3-R51",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IA",
        "ref": "IA-4",
        "ref_name": "Identifier Management",
        "control_text": "c. Assigning the identifier to the intended individual, group, role, service, or device; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R52",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IA",
        "ref": "IA-4",
        "ref_name": "Identifier Management",
        "control_text": "d. Preventing reuse of identifiers for [Assignment: organization-defined time period].",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R53",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-1",
        "ref_name": "Incident Response Policy and Procedures",
        "control_text": "a. Develop, document, and disseminate to [Assignment: organization-defined personnel or \nroles]:\n1. [Selection (one or more): Organization-level; Mission/business process-level; System level] incident response policy that:",
        "objective_text": "(a) Addresses purpose, scope, roles, responsibilities, management commitment, \ncoordination among organizational entities, and compliance; and",
        "examine": ""
    },
    {
        "id": "S3-R54",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-1",
        "ref_name": "Incident Response Policy and Procedures",
        "control_text": "",
        "objective_text": "(b) Is consistent with applicable laws, executive orders, directives, regulations, policies, \nstandards, and guidelines; and",
        "examine": ""
    },
    {
        "id": "S3-R55",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-1",
        "ref_name": "Incident Response Policy and Procedures",
        "control_text": "2. Procedures to facilitate the implementation of the incident response policy and the \nassociated incident response controls",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R56",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-1",
        "ref_name": "Incident Response Policy and Procedures",
        "control_text": "b. Designate an [Assignment: organization-defined official] to manage the development, \ndocumentation, and dissemination of the incident response policy and procedures; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R57",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-1",
        "ref_name": "Incident Response Policy and Procedures",
        "control_text": "c. Review and update the current incident response:\n1. Policy [Assignment: organization-defined frequency] and following [Assignment: \norganization-defined events]; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R58",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-1",
        "ref_name": "Incident Response Policy and Procedures",
        "control_text": "2. Procedures [Assignment: organization-defined frequency] and following [Assignment: \norganization-defined events].",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R59",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-2",
        "ref_name": "Incident Response Training",
        "control_text": "a. Provide incident response training to system users consistent with assigned roles and responsibilities:",
        "objective_text": "1. Within [Assignment: organization-defined time period] of assuming an incident response \nrole or responsibility or acquiring system access",
        "examine": "Is Incident Response Training conducted / accomplished?"
    },
    {
        "id": "S3-R60",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-2",
        "ref_name": "Incident Response Training",
        "control_text": "",
        "objective_text": "2. When required by system changes; and",
        "examine": "When (if any) is training required / mandated?"
    },
    {
        "id": "S3-R61",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-2",
        "ref_name": "Incident Response Training",
        "control_text": "",
        "objective_text": "3. [Assignment: organization-defined frequency] thereafter; and",
        "examine": "When is training reviewed / updated?"
    },
    {
        "id": "S3-R62",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-2",
        "ref_name": "Incident Response Training",
        "control_text": "b. Review and update incident response training content [Assignment: organization-defined \nfrequency] and following [Assignment: organization-defined events].",
        "objective_text": "",
        "examine": "When was training last reviewed / updated, and approved by Executive Management?"
    },
    {
        "id": "S3-R63",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "a. Develop an incident response plan that:\n1. Provides the organization with a roadmap for implementing its incident response \ncapability;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R64",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "2. Describes the structure and organization of the incident response capability;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R65",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "3. Provides a high-level approach for how the incident response capability fits into the \noverall organization;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R66",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "4. Meets the unique requirements of the organization, which relate to mission, size, \nstructure, and functions;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R67",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "5. Defines reportable incidents;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R68",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "6. Provides metrics for measuring the incident response capability within the organization",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R69",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "7. Defines the resources and management support needed to effectively maintain and \nmature an incident response capability;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R70",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "8. Addresses the sharing of incident information;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R71",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "9. Is reviewed and approved by [Assignment: organization-defined personnel or roles]\n[Assignment: organization-defined frequency]; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R72",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "10. Explicitly designates responsibility for incident response to [Assignment: organization defined entities, personnel, or roles]",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R73",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "b. Distribute copies of the incident response plan to [Assignment: organization-defined incident \nresponse personnel (identified by name and/or by role) and organizational elements]",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R74",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "c. Update the incident response plan to address system and organizational changes or \nproblems encountered during plan implementation, execution, or testing;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R75",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "d. Communicate incident response plan changes to [Assignment: organization-defined incident \nresponse personnel (identified by name and/or by role) and organizational elements]; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R76",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "IR",
        "ref": "IR-8",
        "ref_name": "Incident Response Plan",
        "control_text": "e. Protect the incident response plan from unauthorized disclosure and modification.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R77",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-2",
        "ref_name": "Physical Access Authorizations",
        "control_text": "a. Develop, approve, and maintain a list of individuals with authorized access to the facility where the system resides;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R78",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-2",
        "ref_name": "Physical Access Authorizations",
        "control_text": "b. Issue authorization credentials for facility access;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R79",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-2",
        "ref_name": "Physical Access Authorizations",
        "control_text": "c. Review the access list detailing authorized facility access by individuals [Assignment: organization-defined frequency]; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R80",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-2",
        "ref_name": "Physical Access Authorizations",
        "control_text": "d. Remove individuals from the facility access list when access is no longer required.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R81",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-3",
        "ref_name": "Physical Access Control",
        "control_text": "a. Enforce physical access authorizations at [Assignment: organization-defined entry and exit points to the facility where the system resides] by:\n1. Verifying individual access authorizations before granting access to the facility; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R82",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-3",
        "ref_name": "Physical Access Control",
        "control_text": "2. Controlling ingress and egress to the facility using [Selection (one or more): [Assignment: organization-defined physical access control systems or devices]; guards];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R83",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-3",
        "ref_name": "Physical Access Control",
        "control_text": "b. Maintain physical access audit logs for [Assignment: organization-defined entry or exit points];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R84",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-3",
        "ref_name": "Physical Access Control",
        "control_text": "c. Control access to areas within the facility designated as publicly accessible by implementing the following controls: [Assignment: organization-defined physical access controls];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R85",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-3",
        "ref_name": "Physical Access Control",
        "control_text": "d. Escort visitors and control visitor activity [Assignment: organization-defined circumstances requiring visitor escorts and control of visitor activity];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R86",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-3",
        "ref_name": "Physical Access Control",
        "control_text": "e. Secure keys, combinations, and other physical access devices;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R87",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-3",
        "ref_name": "Physical Access Control",
        "control_text": "f. Inventory [Assignment: organization-defined physical access devices] every [Assignment: organization-defined frequency]; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R88",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-3",
        "ref_name": "Physical Access Control",
        "control_text": "g. Change combinations and keys [Assignment: organization-defined frequency] and/or when keys are lost, combinations are compromised, or",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R89",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-6",
        "ref_name": "Monitoring Physical Access",
        "control_text": "a. Monitor physical access to the facility where the system resides to detect and respond to physical security incidents;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R90",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-6",
        "ref_name": "Monitoring Physical Access",
        "control_text": "b. Review physical access logs [Assignment: organization-defined frequency] and upon occurrence of [Assignment: organization-defined events or potential indications of events]; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R91",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PE",
        "ref": "PE-6",
        "ref_name": "Monitoring Physical Access",
        "control_text": "c. Coordinate results of reviews and investigations with the organizational incident response capability.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R92",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PM",
        "ref": "PM-11",
        "ref_name": "Mission And Business Process Definition",
        "control_text": "a. Define organizational mission and business processes with consideration for information security and privacy and the resulting risk to organizational operations, organizational assets, individuals, other organizations, and the Nation; and",
        "objective_text": "",
        "examine": "How are information security and privacy considerations included when evaluating risk to organizational processes? I.e., sensitive information (PII, HIPAA, FCI)."
    },
    {
        "id": "S3-R93",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PM",
        "ref": "PM-11",
        "ref_name": "Mission And Business Process Definition",
        "control_text": "b. Determine information protection and personally identifiable information processing needs arising from the defined mission and business processes; and",
        "objective_text": "",
        "examine": "Was a Data Flow-type analysis accomplished, and are protection measures accounted for and addressed to protect sensitive information?"
    },
    {
        "id": "S3-R94",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PM",
        "ref": "PM-11",
        "ref_name": "Mission And Business Process Definition",
        "control_text": "c. Review and revise the mission and business processes [Assignment: organization-defined frequency].",
        "objective_text": "",
        "examine": "Last review of business processes and their alignment with protecting sensitive information."
    },
    {
        "id": "S3-R95",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-1",
        "ref_name": "Policy And Procedures",
        "control_text": "a. Develop, document, and disseminate to [Assignment: organization-defined personnel or roles]\n1. [Selection (one or more): Organization-level; Mission/business process-level; System level] personnel security policy that:",
        "objective_text": "(a) Addresses purpose, scope, roles, responsibilities, management commitment, coordination among organizational entities, and compliance; and",
        "examine": "Is Personnel Security -type policy or similar business documentation utilized?"
    },
    {
        "id": "S3-R96",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-1",
        "ref_name": "Policy And Procedures",
        "control_text": "",
        "objective_text": "(b) Is consistent with applicable laws, executive orders, directives, regulations, policies, \nstandards, and guidelines; and",
        "examine": "Is policy abiding to industry best practices?"
    },
    {
        "id": "S3-R97",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-1",
        "ref_name": "Policy And Procedures",
        "control_text": "2. Procedures to facilitate the implementation of the personnel security policy and the \nassociated personnel security controls;",
        "objective_text": "",
        "examine": "Are Personnel Security -type procedures or similar business documentation utilized?"
    },
    {
        "id": "S3-R98",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-1",
        "ref_name": "Policy And Procedures",
        "control_text": "b. Designate an [Assignment: organization-defined official] to manage the development, \ndocumentation, and dissemination of the personnel security policy and procedures; and",
        "objective_text": "",
        "examine": "Is policy reviewed / approved for use by executive management or delegate?"
    },
    {
        "id": "S3-R99",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-1",
        "ref_name": "Policy And Procedures",
        "control_text": "c. Review and update the current personnel security:\n1. Policy [Assignment: organization-defined frequency] and following [Assignment: \norganization-defined events]; and",
        "objective_text": "",
        "examine": "When was the last policy update / change approved and published?"
    },
    {
        "id": "S3-R100",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-1",
        "ref_name": "Policy And Procedures",
        "control_text": "2. Procedures [Assignment: organization-defined frequency] and following [Assignment: \norganization-defined events].",
        "objective_text": "",
        "examine": "When was the last procedures update / change approved and published?"
    },
    {
        "id": "S3-R101",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-2",
        "ref_name": "Position Risk Designation",
        "control_text": "a. Assign a risk designation to all organizational positions",
        "objective_text": "",
        "examine": "Are risk designations utilized / assigned to job positions?"
    },
    {
        "id": "S3-R102",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-2",
        "ref_name": "Position Risk Designation",
        "control_text": "b. Establish screening criteria for individuals filling those positions; and",
        "objective_text": "",
        "examine": "Are screening / background checks / interviews accomplished prior to onboarding personnel?"
    },
    {
        "id": "S3-R103",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-2",
        "ref_name": "Position Risk Designation",
        "control_text": "c. Review and update position risk designations [Assignment: organization-defined frequency",
        "objective_text": "",
        "examine": "Are risk designations reviewed / approved?"
    },
    {
        "id": "S3-R104",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-3",
        "ref_name": "Personnel Screening",
        "control_text": "a. Screen individuals prior to authorizing access to the system; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R105",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-3",
        "ref_name": "Personnel Screening",
        "control_text": "b. Rescreen individuals in accordance with [Assignment: organization-defined conditions \nrequiring rescreening and, where rescreening is so indicated, the frequency of rescreening].",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R106",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-4",
        "ref_name": "Personnel Termination",
        "control_text": "Upon termination of individual employment:\na. Disable system access within [Assignment: organization-defined time period];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R107",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-4",
        "ref_name": "Personnel Termination",
        "control_text": "b. Terminate or revoke any authenticators and credentials associated with the individual;",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R108",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-4",
        "ref_name": "Personnel Termination",
        "control_text": "c. Conduct exit interviews that include a discussion of [Assignment: organization-defined information security topics];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R109",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-4",
        "ref_name": "Personnel Termination",
        "control_text": "d. Retrieve all security-related organizational system-related property; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R110",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "PS",
        "ref": "PS-4",
        "ref_name": "Personnel Termination",
        "control_text": "e. Retain access to organizational information and systems formerly controlled by terminated individual.",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R111",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "SA",
        "ref": "SA-9",
        "ref_name": "External System Services",
        "control_text": "a. Require that providers of external system services comply with organizational security and \nprivacy requirements and employ the following controls: [Assignment: organization-defined \ncontrols];",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R112",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "SA",
        "ref": "SA-9",
        "ref_name": "External System Services",
        "control_text": "b. Define and document organizational oversight and user roles and responsibilities with regard \nto external system services; and",
        "objective_text": "",
        "examine": ""
    },
    {
        "id": "S3-R113",
        "sheet": "3. Training, HR, IRP, FSO",
        "family": "SA",
        "ref": "SA-9",
        "ref_name": "External System Services",
        "control_text": "c. Employ the following processes, methods, and techniques to monitor control compliance by \nexternal service providers on an ongoing basis: [Assignment: organization-defined processes, \nmethods, and techniques]",
        "objective_text": "",
        "examine": ""
    }
]
