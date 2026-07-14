# ✈️ Interactive Vacation Approval Engine

An automated, serverless vacation approval loop built using Google Forms, Google Apps Script, and Google Sheets, featuring interactive email-action endpoints for rapid management processing.

---

## 🔄 The Interaction Cycle
[Employee Form Submit] ──► [OnFormSubmit Trigger] ──► [Interactive HTML Mail Generated]
│
▼
[Manager Clicks Approved/Denied] ◄─────────────────── [Approve/Deny Buttons]
│
▼
[GET Request Web App Endpoint] ──► [Writes Status to Spreadsheet] ──► [Alerts Employee]

## 🛠️ Highlights
* **One-Click Decisions:** Zero portal logins required; managers can approve or deny requests directly from inside their inbox on any mobile device.
* **Robust Exception Handling:** Coordinates query parameters and status writes on high-concurrency requests.
