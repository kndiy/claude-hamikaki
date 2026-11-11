# GMPc (group)

Rules here apply to every project below this folder. Put project-specific rules in that project's own CLAUDE.md.

## Scope
Data, reports and documents for **GMPc**, a consulting company for Good Manufacturing Practice (GMP) and related standards (GSP/GDP, HS-GMP, cGMP, ISO). Its services:
1. **Consulting packages**: advising customers on meeting GMP requirements.
2. **Permit filing on the customer's behalf**: GMPc staff draft and submit forms and dossiers to the authorities to get permits/certifications for the customer.
3. **Contractor work**: building and setting up the required infrastructure and training the customer's staff to run the factory.

Typical projects: customer and contract statistics, reconciling data against NAS records, drafting dossiers and forms, training material.

## Shared rules
- **NAS drives:** every drive other than `C:` is a GMPc NAS share on `\\gmpc.myds.me`, currently `H:` (Shared), `K:` (Regulations library), `L:` (Sales), `N:` (Accounting - Projects), `O:` (HR/Admin), `Q:` (Projects). Read or copy only, into the local project (`data/raw/`). Never edit, rename or delete anything there. Record the exact source path in the project's CLAUDE.md. On another machine or inside WSL, confirm the NAS path before using it. (This rule moved here from the root CLAUDE.md on 2026-11-11, because only GMPc uses the NAS.)
- **Customer data is confidential** (contract values, contacts, personal data). Keep it in the local project, and never send it to external services or publish it unless the user asks.
- **Dossiers submitted to authorities must be truthful.** Personnel named in a dossier must be people who actually hold that role for the customer. If GMPc staff fill a role, it must be through a real, documented secondment or employment arrangement. Claude won't produce dossiers that present people in roles they don't actually hold.
