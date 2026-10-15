---
title: Flattening case navigation to remove nested tabs
description: How we redesigned the information architecture inside a case so prosecutors can reach charges, defendants and witnesses without working through up to 4 layers of navigation.
---

<!-- DRAFT — working post, updated as the work progresses. Last updated 14 October 2026. -->

Inside a case in the Digital Case File (DCF), information sits behind up to 4 layers of navigation: secondary navigation, then tabs, then sub-tabs, then navigation inside individual cards. Users told us the most important information, especially charges, is buried. We are redesigning the information architecture (IA) inside a case to remove nested tabs. The work so far is based on a team workshop on screenshots of the live service, an early tree test and 2 moderated interviews.

## Where we started

A case has 4 secondary navigation items: Details, Materials, Indictments and Reviews.

Details then has its own [tabs](https://design-system.service.gov.uk/components/tabs/): Overview, Outline, Defendants, Witnesses and History. Several of these have more navigation inside them:

- Outline has 4 sub-tabs: Factual summary, Confidential information to prosecutor, Certifications, and Rationale for unauthorised changes
- each defendant card has its own sub-navigation: Charges, Defendant details, Bail and remand, and Candidate court orders
- each witness has tabs for their details and statements

Materials and Reviews also use tabs inside the page.

<!-- SCREENSHOT PLACEHOLDER — duplicate as needed -->
![Screenshot described under image](current-case-details-tabs.png)

*Caption: The current Details page, with tabs and sub-tabs*

Describe the screenshot here. Include the page heading, what each section contains, and what the buttons and links say.
<!-- END PLACEHOLDER -->

## What we learned

From the workshop and research, we found that:

- charges are the most useful way to understand a case, but they sit 3 layers deep inside each defendant card
- users could not tell the difference between Overview and Outline
- people mistook sub-navigation links for table headers
- users think in cases first, and treat tasks as prompts for action within a case
- people returning to a case need to reorient quickly: what the case is, where it is up to and what has already been done
- monitoring codes, such as rape and serious sexual offences (RASSO), are hard to notice in their current list

As one participant put it, charges "feel buried several tabs in".

## What we did

### Agreed a single layer of navigation inside a case

We replaced the 4 secondary navigation items and the tabs inside Details with 8 secondary navigation items:

1. Overview
2. Tasks
3. Defendants
4. Witnesses
5. Materials
6. Indictments
7. Reviews
8. History

Our aim is to remove every tab from the case pages. Where a page has several parts, it becomes sections on one page or a set of linked pages.

| Before | After |
|---|---|
| Details > Overview tab | Overview |
| Details > Outline tab and its 4 sub-tabs | Sections on Overview |
| Overview > Tasks panel | Tasks |
| Details > Defendants tab, with sub-navigation in each card | Defendants list, then a page for each defendant |
| Details > Witnesses tab | Witnesses list, then a page for each witness |
| Details > History tab | History |

### Merged Overview and Outline

The new Overview brings together case status, a summary of charges across all defendants, key contacts, monitoring codes, and all 4 parts of the old Outline. This answers "what is this case and where is it up to" on one page, and removes the confusing choice between Overview and Outline.

### Gave defendants and witnesses their own pages

Defendants and Witnesses each get a list page that links to a page for each person. On those pages, sections replace the sub-tabs. We kept the list page even for cases with 1 defendant, so the pattern is consistent.

### Made tasks part of the case navigation

Tasks inside a case are the same tasks as in the main Tasks list, filtered to that case.

### Produced a sitemap

We captured the new structure as a sitemap, with the reason for each change and the evidence behind it, to review with the team.

<!-- SCREENSHOT PLACEHOLDER — duplicate as needed -->
![Screenshot described under image](case-sitemap.png)

*Caption: The proposed sitemap for pages inside a case*

Describe the screenshot here. Include the page heading, what each section contains, and what the buttons and links say.
<!-- END PLACEHOLDER -->

## What we have not resolved yet

- **Materials, Indictments and Reviews:** their internal structure is out of scope, but Materials and Reviews still use tabs that will need the same treatment.
- **Offences taken into consideration and associated cases:** these appear in an earlier proposed design but not in the current service, so we need to decide whether they are in scope.
- **History:** this currently shows only changes to the factual summary, but users need a full case chronology.
- **Placement on Overview:** we have not decided whether a short task summary stays on Overview, or whether the under-18 warning moves into a persistent case status strip.
- **Section order and in-page navigation:** we still need to decide how users move around long pages, for example with anchor links or side navigation.
- **Testing:** our evidence comes from small samples, so we need to test the new structure, for example with a tree test.
