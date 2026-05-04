# KICKR RUN CORE (DW2171) — Risk Register
**Project:** KICKR RUN CORE  
**Last Updated:** 2026-03-26  
**Target Ship Date:** 2027-03-06  
**Owner:** PM — Wahoo Fitness  

---

## Risk Rating Key

| Probability | Impact | Risk Score |
|---|---|---|
| H = High (>70%) | H = High (schedule slip >2 weeks, cost impact, compliance failure) | Critical = H/H |
| M = Medium (30–70%) | M = Medium (1–2 week slip, manageable cost) | High = H/M or M/H |
| L = Low (<30%) | L = Low (<1 week slip, minimal cost) | Medium = M/M |
| | | Low = L/L, L/M, M/L |

---

## Risk Register

| ID | Risk Title | Description | Category | Probability | Impact | Rating | Early Warning Signals | Mitigation Actions | Owner | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| R-01 | MDU (BQ) Delivery Slip | BQ MDU sample not ready by June 30, 2026. Updated comm spec was delivered late (Feb 25). Current tracking already at 125 days vs 118 planned. PM notes explicitly flag this timeline as unacceptable. | Schedule / Supplier | H | H | **Critical** | BQ misses any intermediate milestone; no weekly progress update provided | (1) Demand milestone plan from BQ with weekly check-ins. (2) Discuss local resource augmentation with Bisser. (3) Formally notify EB that June 30 is the hard ceiling and document the risk if BQ cannot commit. | PM + Bisser | Open |
| R-02 | Zero Schedule Float in Back Half | No contingency exists between any major milestones from July 2026 onward. EVT→Tooling→DVT→PVT→MP are all back-to-back with 0 buffer days. Any single slip cascades directly to ship date. | Schedule | H | H | **Critical** | Any milestone in the chain slips by even 1 day | (1) Insert a minimum 2-week buffer between EVT confirmation and tooling release. (2) Present revised schedule with honest ship date to Wahoo leadership. (3) Identify which downstream phases have any flexibility (e.g., DVT production duration). | PM | Open |
| R-03 | Historical Overrun Pattern | Every completed phase has exceeded its planned duration. Initial EVT sample: 30 days planned, 48 actual (60% overrun). MDU: 118 planned, 125+ actual. Applying this pattern to remaining phases puts ship date at risk of 3–6 week slip. | Schedule | H | H | **Critical** | Remaining phase durations tracking above plan at midpoint | (1) Reforecast remaining phases using actuals-based estimates, not original plan. (2) Use 48-day EVT overrun as a reference data point when negotiating tooling and DVT timelines with EB. | PM | Open |
| R-04 | CAD Approval Window Too Narrow | EB must deliver updated 3D files by Apr 18; Wahoo must approve by Apr 24 (6-day review window). Any internal delay or revision request compresses the mechanical sample manufacturing window (Apr 25 – May 10) and risks the planned factory visit (May 10–15). | Schedule / Engineering | M | H | **High** | Wahoo reviewer unavailable; revision request submitted after Apr 20 | (1) Confirm Wahoo internal reviewer is assigned and available the week of Apr 21. (2) Pre-align on known open items before EB delivers files to reduce revision likelihood. (3) If files slip past Apr 18, immediately assess impact to factory visit dates. | PM + Engineering | Open |
| R-05 | Factory Visit Timing Conflict | Wahoo visit to EB planned for May 10–15, which is the last day of mechanical sample manufacturing. Samples may not be complete or validated in time for a meaningful visit. | Schedule / Engineering | M | M | **Medium** | Mechanical samples not assembled by May 8 | (1) Confirm with EB that samples will be ready for review by May 9 at the latest. (2) Define a clear agenda and acceptance criteria for the visit now so scope is bounded. (3) Prepare contingency — if samples aren't ready, convert visit to firmware/integration review only. | PM + EB | Open |
| R-06 | Regulatory Certification Not Tracked | No FCC, CE, or market certification milestones are visible in the current schedule. EN957 is referenced in DVT testing but submission timelines, lab bookings, and sample requirements are not mapped. Certification can take 8–12 weeks and requires locked hardware. | Compliance / Schedule | M | H | **High** | DVT sample completion without a lab booking confirmed | (1) Immediately map required certifications by target market (FCC, CE, EN957, others). (2) Identify earliest DVT sample availability for lab submission. (3) Add certification milestones to master schedule. (4) Book lab slots now — labs often have 4–6 week lead times. | PM + Compliance | Open |
| R-07 | Communication Protocol Spec Delivered Late | Updated MDU comm spec was not delivered to EB/BQ until Feb 25, 2026 — well into the EVT phase. This directly caused the MDU delay. Risk is that further spec changes are made during EVT or early DVT, resetting BQ development work. | Engineering / Scope | M | H | **High** | Any Wahoo-side discussion of protocol changes after Apr 2026 | (1) Formally freeze the MDU communication spec and document the freeze date. (2) Establish a change control process — any spec change requires PM approval and schedule impact assessment before EB/BQ are informed. | PM + Engineering | Open |
| R-08 | CAD Format Conversion Delay | Wahoo requested conversion from GBL to OBJ format on Mar 24. If EB's conversion introduces errors or requires re-validation, it could delay the Apr 18 3D file delivery. | Engineering | L | M | **Low** | EB flags conversion issues before Apr 10 | (1) Request confirmation from EB that OBJ conversion is underway and on track. (2) Set Apr 10 as internal check-in date on conversion status — provides 8 days to course correct before the Apr 18 deadline. | PM + EB | Open |
| R-09 | CNY Holiday Impact on MP Production | MP production window (Dec 26, 2026 – Feb 25, 2027) overlaps with Chinese New Year. Schedule notes 20 days of holiday within a 61-day window, reducing actual working days to ~40. Any pre-CNY slip in PVT reduces the production buffer before the holiday. | Schedule / Manufacturing | M | M | **Medium** | PVT production start slipping past Dec 1, 2026 | (1) Confirm CNY dates for 2027 and validate the 20-day assumption. (2) Identify which MP tasks must be completed before CNY shutdown. (3) Work with EB to pre-stage materials and tooling so production can resume immediately after the holiday. | PM + EB Manufacturing | Open |
| R-10 | DVT Testing Duration Fixed and Non-Compressible | Life/EN957/durability testing is 60–65 days (Sept 16 – Nov 20) and cannot be accelerated. This is a hard constraint. Any slip entering DVT testing hits PVT start date directly with no ability to recover time. | Schedule / Testing | H | H | **Critical** (if upstream slips) | Any milestone upstream of Sept 16 DVT test start slipping | (1) Treat Sept 16 as a hard backstop when managing all upstream dates. (2) Identify whether any DVT testing can begin in parallel with late-stage DVT production (e.g., safety testing on pre-production units). (3) Pre-book test lab or internal test capacity now. | PM + Test Engineering | Open |

---

## Open Action Items from Risk Register

| ID | Action | Due Date | Owner |
|---|---|---|---|
| A-01 | Obtain milestone plan from BQ with weekly check-in cadence | 2026-04-04 | PM + Bisser |
| A-02 | Discuss local resource augmentation to accelerate BQ MDU work | 2026-04-04 | PM + Bisser |
| A-03 | Formally notify EB that June 30 MDU date is the hard ceiling | 2026-04-04 | PM |
| A-04 | Reforecast schedule using actuals-based estimates and present to leadership | 2026-04-10 | PM |
| A-05 | Confirm Wahoo CAD reviewer is assigned and available week of Apr 21 | 2026-04-07 | PM |
| A-06 | Map all required regulatory certifications and add to master schedule | 2026-04-10 | PM + Compliance |
| A-07 | Book certification lab slots | 2026-04-10 | Compliance |
| A-08 | Formally freeze MDU communication spec and document change control process | 2026-04-07 | PM + Engineering |
| A-09 | Request OBJ conversion status update from EB | 2026-04-04 | PM |
| A-10 | Confirm CNY 2027 dates and validate MP production assumptions | 2026-04-10 | PM |
| A-11 | Define agenda and acceptance criteria for May factory visit | 2026-04-10 | PM |

---

*Generated from project schedule and PM notes dated 2026-03-26.*
