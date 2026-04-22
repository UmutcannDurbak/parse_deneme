# PROJECT_MEMORY.md
> **Project Name:** Parse Deneme
> **Last Updated:** 2026-04-22
> **Current Phase:** Maintenance / Feature Addition
> **Active Context:** Added new Bask Cheesecake variants, Trileçe quantity shift, and EKŞİ MAYALI KÖY EKMEĞİ support.

---

## [1. PROJECT VISION & GOALS]
* **Core Concept:** Automating parsing and shipment operations for products (tatlı, cheesecake, etc.).
* **Target Audience:** Internal operations team.
* **Success Criteria:** Correct listing of all products across different categories.

## [2. TECH STACK & CONSTRAINTS]
* **Language/Framework:** Python
* **Backend/DB:** Excel/CSV files (sevkiyat_donuk, sevkiyat_tatlı, etc.)
* **State Management:** N/A
* **Key Packages:** pandas, openpyxl
* **Constraints:** Maintain consistency with older working versions for non-cheesecake products.

## [3. ARCHITECTURE & PATTERNS]
* **Design Pattern:** Functional and OOP mix (shipment_oop.py suggests OOP).
* **Folder Structure:** Root directory contains script files and data files.
* **Naming Conventions:** snake_case for scripts and variables.

## [4. ACTIVE RULES (The "Laws")]
1.  Follow the specific rollback request for `parse_gptfix.py`, `shipment_oop.py`, and `tatli_siparis.py`.
2.  Maintain `PROJECT_MEMORY.md`.
3.  Ensure non-cheesecake products are listed correctly.

## [5. PROGRESS & ROADMAP]
- [x] Phase 1: Setup & Configuration
- [x] Phase 2: Core Features
    - [x] Initial development
    - [x] Bug fixing (Product listing issue via Rollback)
- [x] Phase 3: Rollback to v1.3.41 for specific files
- [ ] Phase 4: UI Polish (if any)
- [ ] Phase 5: Testing & Deployment

## [6. DECISION LOG & ANTI-PATTERNS]
* **[Karar - 2026-04-22]:** Replaced old cheesecake (SEBASTIAN, FRAMBUAZ) logic with new Bask variants (SADE, MATCHA, YABAN MERSİNLİ).
* **[Logic Change - 2026-04-22]:** Reverted TRİLEÇE quantity placement to "same cell" as per user clarification.
* **[Karar - 2026-04-22]:** Added support for "EKŞİ MAYALI KÖY EKMEĞİ". The quantity is written in the same cell as the name with a "KL." unit fixed.

---
**OPERATIONAL DIRECTIVE:**
1.  **Read First:** Before answering any prompt, check this file for context.
2.  **Update Often:** If a task is completed, check the box [x]. If a tech decision changes, update Section 2.
3.  **Stay Consistent:** Do not suggest code that violates "Active Rules".
