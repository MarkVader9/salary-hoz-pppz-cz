# SALARY s.r.o. - Czech Health Insurance Interfaces (XSD/XML)

<!-- ISO 9001:2015 Document Reference: REF-IT-XSD-01 | Version: 1.0.0 | Status: APPROVED -->

Master repository containing official XSD schemas and XML templates for regulatory reporting to Czech Health Insurance Companies (Zdravotní pojišťovny ČR). Developed and maintained by **SALARY s.r.o.** as an integral part of our payroll core software validation platform.

## 📌 Project Overview

This repository acts as the single source of truth (SSOT) for data validation formats required by corporate payroll processing. It ensures strict adherence to Czech statutory requirements and **ISO 9001:2015 Quality Management Standards** (specifically Clause 8.1 - Operational Planning and Control).

### Managed Document Interfaces:
*   **HOZ** (Hromadné oznámení zaměstnavatele) - Employee registration, health insurance company changes, and category modifications (Statutory limit: 8 days).
*   **PPPZ** (Přehled o platbě pojistného zaměstnavatele) - Monthly total premium summary, calculation bases, and employer declarations (Statutory limit: 20th day of the following month).

## 🗂 Repository Structure

```text
├── .github/              # CI/CD pipelines for XML automated validation
├── schemas/              # Official XSD files
│   ├── hoz/
│   │   └── hoz_v2026.xsd # Schema for Hromadné oznámení zaměstnavatele
│   └── pppz/
│       └── pppz_v2026.xsd # Schema for Přehled o platbě pojistného
├── templates/            # Production-ready XML payload samples
│   ├── hoz_sample.xml
│   └── pppz_sample.xml
└── README.md             # This documentation
```

## ⚖️ Legal & Regulatory Framework

All schemas in this repository comply with current Czech Republic legislation:
*   **Act No. 48/1997 Coll.** on Public Health Insurance
*   **Act No. 592/1992 Coll.** on Premiums for Public Health Insurance
*   **Act No. 300/2008 Coll.** on Electronic Acts and Authorized Document Conversion (Data Box delivery protocol)

## 🛠 Integration & CI/CD Validation

Every commit to the `main` branch triggers an automated workflow that runs schema compliance checks.

### Quick Manual Validation (CLI)
To validate an XML export against our production XSD locally, run:
```bash
xmllint --schema schemas/hoz/hoz_v2026.xsd templates/hoz_sample.xml --noout
```

## 🔒 Security & Data Privacy (GDPR)
*   **Zero PII Policy:** This repository strictly forbids the inclusion of real Personal Identifiable Information (PII). 
*   All data files under `templates/` use randomly generated mock variables.
*   Access tokens and electronic signature certificates must be configured via environment secrets (`GH_SECRETS`) and never hardcoded.

## 👥 Authors & Change Management

*   **Maintainer:** SALARY s.r.o. Development Team (<dev@salary.cz>)
*   **Process Owner:** Compliance & Quality Assurance Manager
*   **ISO 9001 Audit Trail:** Any alteration to `.xsd` definitions requires a formal Change Request (CR) protocol log.

---
© 2026 SALARY s.r.o. All rights reserved. Internal Technical Documentation.
