# M365 Pulse Report Builder

> **Privacy-first, standalone security posture analyzer and report generator for Microsoft 365.**  
> Transforms M365 Pulse assessment `.xlsx` workbooks into interactive, accessible, executive-ready HTML reports.

[![Platform](https://img.shields.io/badge/Platform-Modern%20Browser%20%7C%20Node.js%2018%2B-blue.svg)](#requirements)
[![Privacy](https://img.shields.io/badge/Privacy-100%25%20Offline%20%26%20Local-success.svg)](#privacy--security-by-design)
[![Dependencies](https://img.shields.io/badge/Dependencies-Zero-brightgreen.svg)](#technical-architecture)
[![Accessibility](https://img.shields.io/badge/Accessibility-WCAG%20AA%20%2F%20AAA-orange.svg)](#accessible--high-contrast-design)

---

## Highlights

- **100% Offline & Private:** No network requests, no external telemetry, no cloud upload, no AI hallucinations. Your sensitive tenant data and account details never leave your computer.
- **Zero Dependencies:** Pure vanilla JavaScript and HTML. Built-in, zero-dependency streaming `.xlsx` unpacker and parser.
- **Resilient Signature Detection:** Sheets are identified by column header patterns, not brittle sheet names. Works across various PowerShell export script versions, reordered sheets, and naming conventions.
- **Accessible & High-Contrast Design:** Engineered for readability with WCAG AAA compliant contrast, explicit text badges (colorblind-friendly), and no low-contrast pastel colors.
- **Interactive Report Output:** Built-in live search, real-time severity filters, sticky section navigation, and clean print-to-PDF formatting.

---

## Security Domains Covered

The engine analyzes up to 81 workbook sheets and correlates posture data across core Microsoft 365 and Entra ID domains:

| Domain | Key Insights & Security Controls |
| :--- | :--- |
| **Executive Posture & Secure Score** | 90-day Secure Score history, trend sparklines, point gap analysis, workload distribution (Core M365, Defender for Endpoint, Identity, Cloud Apps), and top quick-win recommendations. |
| **Privileged Access & PIM** | Standing permanent directory role assignments vs. Just-In-Time (JIT) eligible roles, over-privileged accounts holding excessive roles, and Global Administrator count auditing. |
| **Identity & Authentication** | Tenant-wide MFA registration percentage, admin accounts using weak MFA (SMS or voice), dormant/stale accounts, and Identity Protection risky user detection. |
| **Conditional Access (CA)** | Policy enforcement state (`On`, `Report-Only`, `Off`), break-glass emergency account exclusions, empty target validations, and authentication strength requirements. |
| **Endpoint & Device Security (MDM)** | BYOD personal device enrollment restrictions (Windows, iOS, macOS, Android), Windows LAPS status, device registration MFA requirements, and unmanaged/registered device inventory. |
| **SharePoint & OneDrive Governance** | Legacy authentication protocols, external guest re-sharing permissions, idle session timeouts, and domain sharing restrictions. |
| **App Registrations & Credentials** | Expiring or expired client secrets and certificates, ownerless application registrations, and multi-tenant app exposure. |
| **Tenant Setting Checks** | Automatic ingestion of pass/fail baseline checks across tenant authorization policies, admin consent workflows, password protection, and security defaults. |
| **License Matrix** | Subscribed vs. assigned licenses across Microsoft 365, Entra ID P1/P2, and add-on SKUs. |

---

## How to Use

### 1. Web Browser (Drag & Drop)

No installation or web server required:

1. Open `index.html` in Microsoft Edge, Google Chrome, Brave, or Firefox.
2. Drag and drop your `M365-Pulse-Report.xlsx` file onto the upload target (or click to browse).
3. The report is instantly parsed and rendered in an interactive preview pane.
4. Click **Save HTML report** to download the standalone report or **Open in new tab** to view it full-screen.

### 2. Command Line (Node.js 18+)

Ideal for scheduled audits, batch evaluations, and CI/CD reporting pipelines:

```bash
# Generate report.html next to the input workbook
node cli.js M365-Pulse-Report.xlsx

# Specify custom output path
node cli.js M365-Pulse-Report.xlsx -o C:\Reports\Apexogy-Pulse.html

# Batch process multiple workbooks into a target folder
node cli.js C:\Assessments\*.xlsx -d C:\Reports
```

The CLI prints a tenant summary, detected mappings, and finding counts directly to the console:

```text
M365-Pulse-Report.xlsx -> M365-Pulse-Report.html
  Tenant: Apexogy   Findings: 5 critical, 13 high, 10 medium
  ok  Secure Score history           00_Tenant_Score
  ok  Secure Score controls          00_Tenant_Risks
  ok  Privileged Roles (PIM)         43_Entra_PIM_Permanent_Users
  ok  Device enrollment restrictions 42_DeviceEnrollment_MDM
  ok  SharePoint & OneDrive settings 46_Tenant_SPOSettings
  ...
```

---

## Interactive Report Features

The generated HTML report is completely standalone (CSS, SVG icons, and vanilla JavaScript are bundled inline):

- **Sticky Navigation Bar:** Quickly jump between report sections (`Overview`, `Findings`, `PIM & Roles`, `Devices & MDM`, `SharePoint & Collab`, `Identity & MFA`, `Conditional Access`, `Apps`, `Setting Checks`).
- **Interactive Severity Filters:** Filter findings by `All`, `Critical`, `High`, `Medium`, or `Info` with real-time counters.
- **Live Search:** Instant client-side filtering across findings and tenant setting checks by typing keywords (e.g., *MFA*, *Guest*, *LAPS*, *SharePoint*).
- **Accessible Account Badges:** User chips feature explicit text badges (`[MFA OK]`, `[NO MFA]`, `[ADMIN]`, `[SERVICE]`, `[GUEST]`) rather than color-alone indicators, ensuring clarity for colorblind users.
- **Print & PDF Optimization:** Pre-configured `@media print` CSS cleanly strips navigation bars, formats tables, and applies `page-break-inside: avoid` for professional PDF export via browser print (`Ctrl+P`).

---

## How Sheet Detection Works

Sheet names are **never hardcoded**. Instead, `pulse-engine.js` inspects the headers of each sheet against column signature rules:

1. **Normalized Aliases:** Headers are compared in lowercase with whitespace and special characters stripped (e.g., `MFA Registered`, `MFARegistered`, and `mfa_registered` all normalize to `mfaregistered`).
2. **Signature Groups:** A sheet is matched to a security domain when it satisfies all required column groups defined in `SIGNATURES`.
3. **Generic Settings Check:** Any unclaimed sheet with $\le 40$ rows and a `Status`, `Result`, or `Compliant` column is automatically ingested into the **Tenant Setting Checks** table.

### Primary Sheet Signatures

| Signature Key | Required Column Signatures | Target Data |
| :--- | :--- | :--- |
| `pimPerm` | `rolename`, `userprincipalname` / `memberdisplayname`, `membertype` | Permanent directory role assignments |
| `pimElig` | `rolename`, `userprincipalname` / `principaldisplayname`, `membertype` | PIM eligible role assignments |
| `deviceMDM` | `windowspersonaldeviceenrollmentblocked`, `iospersonaldeviceenrollmentblocked` | Intune personal device enrollment restrictions |
| `devicePolicy` | `azureadregistrationallowed`, `lapsenabled` / `lapsstatus` | Device registration policy & Windows LAPS |
| `deviceConfig` | `operatingsystem`, `jointype` | Joined & registered device inventory |
| `spoSett` | `islegacyauthprotocolsenabled`, `isresharingbyexternalusersenabled` | SharePoint Online external sharing & legacy auth |
| `users` | `upn` / `userprincipalname`, `mfaregistered` / `mfastatus` | User accounts, MFA methods, sign-in freshness |
| `cap` | `policyname` / `displayname`, `state`, `grantcontrols` | Conditional Access policies |
| `apps` | `applicationname` / `displayname`, `ownercount` / `owners`, `secretdaysleft` / `certdaysleft` | App registrations & credential expiry |
| `score` | `currentscore`, `maxscore`, `date` / `createddatetime` | Secure Score historical progression |

---

## Customization & Tuning

All configuration, signature rules, and finding criteria are centralized in [`pulse-engine.js`](file:///D:/Git/CC/M365-Pulse-Report-App/pulse-engine.js):

### Adjusting Thresholds (`CONFIG`)

```javascript
const CONFIG = {
  staleDays: 90,            // Inactive account threshold (days)
  expiryCriticalDays: 7,    // App credential expiry critical threshold
  expiryWarnDays: 30,       // App credential expiry warning threshold
  mfaCriticalBelow: 50,     // Critical alert threshold for tenant MFA %
  mfaHighBelow: 90,         // High alert threshold for tenant MFA %
  maxGlobalAdmins: 5,       // Maximum recommended permanent Global Admins
  maxCheckSheetRows: 40,    // Maximum rows for generic pass/fail check sheets
  topRiskControls: 8,       // Number of top controls shown in quick wins
};
```

### Adding New Finding Rules (`analyse()`)

Inside `analyse(wb, fileName)`, you can add custom security rules using the `add()` helper:

```javascript
add(severity, area, title, detail, sourceSheetName);
// severity: 'crit' | 'high' | 'med' | 'info'
```

### Adding New Column Aliases (`SIGNATURES`)

If an exporter uses an unlisted column header, add the normalized alias (lowercase, alphanumeric only) to the corresponding signature array in `SIGNATURES`:

```javascript
admins: { 
  need: [
    ['userprincipalname', 'upn', 'memberdisplayname'], 
    ['adminroles', 'roles', 'directoryroles', 'rolename']
  ] 
}
```

---

## File Structure

```text
M365-Pulse-Report-App/
├── cli.js               # CLI runner for Node.js batch and automated execution
├── index.html           # Standalone browser drag-and-drop web interface
├── pulse-engine.js      # Core engine: xlsx parser, posture analysis, HTML renderer
├── README.md            # Documentation and usage guide
└── M365-Pulse-Report.xlsx # Example assessment workbook input
```

---

## Technical Architecture

- **Format:** Standalone HTML5 with responsive CSS Grid/Flexbox layouts.
- **Fonts:** Loads Google Fonts (*IBM Plex Sans*, *Bricolage Grotesque*, *IBM Plex Mono*) when connected to the internet; automatically falls back to native system fonts (`Segoe UI`, `system-ui`, `ui-sans-serif`) when offline.
- **Compatibility:** Edge, Chrome, Safari, Firefox, and Node.js 18+.

---

## License

Internal Microsoft 365 Assessment Tool. Distributed for security posture assessment and auditing purposes.
