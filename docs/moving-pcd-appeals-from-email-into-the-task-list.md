---
title: Moving PCD appeals from email into the task list
description: How we designed a task-based way for Charging Managers and District Crown Prosecutors to triage, assign and review police appeals against charging decisions, informed by research with 9 users and product owner decisions for MVP.
---

When a CPS lawyer decides not to charge a suspect, the police can appeal that pre-charge decision (PCD). Today these appeals arrive by email, with no task, reminder or audit trail, so they can go unactioned without anyone noticing. We designed a way to manage them as tasks in the case management system (CMS). The work was informed by 9 research sessions, a draft to-be process and decisions with the product owner on what goes into the minimum viable product (MVP).

## Where we started

An appeal is a police officer, usually an inspector or above, challenging a decision not to charge. It should arrive within 14 days of the original decision.

At the moment:

- appeals arrive in a shared inbox or Teams channel
- a Charging Manager adds the email to the case in CMS, creates a task and assigns it to a District Crown Prosecutor (DCP)
- in CPS Direct (CPSD), the out-of-hours charging service, appeals are posted in a Teams channel for any available DCP to pick up
- each area tracks appeals in its own spreadsheet

There are 2 kinds of appeal. Most are routine. Priority appeals, also called red cases, usually involve someone in custody, so they are urgent straight away.

## What we learned

We ran 9 remote sessions with 4 Charging Managers and 5 DCPs, including people from CPSD and a specialist rape and serious sexual offences (RASSO) unit. Each session explored the current process, then tested a clickable prototype covering viewing, assigning and starting an appeal task.

Participants strongly supported moving to tasks. As one DCP put it: "It doesn't appear in a task list, it doesn't appear in a reminder or any other format. It's simply an e-mail."

We also found that:

- people weren't sure where to click to start working on an appeal, and asked whether "Show" should say "action" instead
- 2 participants couldn't find the link into case review
- CPSD DCPs pick appeals up from a shared pool rather than being assigned them
- Charging Managers want to see when the original decision was made, so they can check the appeal is in time
- the reason for appeal must stay free text, but long unbroken blocks are hard to read
- DCPs need to know the case has been reactivated and that a new charging decision must be recorded

## What we changed

### Showing both types of appeal in one task list

There are 2 task types: "Review PCD appeal" and "Review priority PCD appeal". Priority appeals keep their own Priority charging tab, but now also appear in the main task list, where users can filter by task type. This matches how the live service will work.

We also changed the prototype to use the official task names, so users see the same words as in the live service.

<!-- SCREENSHOT PLACEHOLDER — duplicate as needed -->
![Screenshot described under image](task-list-both-appeal-types.png)

*Caption: The task list showing routine and priority appeals together*

Describe the screenshot here. Include the page heading, what each section contains, and what the buttons and links say.
<!-- END PLACEHOLDER -->

### Making the next action obvious

We renamed the "Details" column to "Actions", in response to the confusion about where to click. Expanding a row now shows the same 3 [buttons](https://design-system.service.gov.uk/components/button/) on every appeal:

- View appeal details (secondary)
- Reassign task (secondary)
- Start task (primary, on the right)

Previously, unassigned tasks showed a different "Assign" button. Using the same 3 actions on every row means users don't have to work out why the options change.

<!-- SCREENSHOT PLACEHOLDER — duplicate as needed -->
![Screenshot described under image](task-list-expanded-row.png)

*Caption: An expanded task list row with its 3 actions*

Describe the screenshot here. Include the page heading, what each section contains, and what the buttons and links say.
<!-- END PLACEHOLDER -->

### Separating viewing an appeal from starting it

Starting the task takes the DCP into the appeal page, with a button to continue to case review. We added a read-only version of the same page, without that button.

This lets a Charging Manager check an appeal before assigning it, and lets anyone read it without being pushed into a review. The back link returns users to wherever they came from.

<!-- SCREENSHOT PLACEHOLDER — duplicate as needed -->
![Screenshot described under image](appeal-read-only.png)

*Caption: The read-only appeal page*

Describe the screenshot here. Include the page heading, what each section contains, and what the buttons and links say.
<!-- END PLACEHOLDER -->

### Showing a snapshot of the appeal on the case overview

The case overview now has a [summary card](https://design-system.service.gov.uk/components/summary-list/#summary-cards) titled with the task name. It shows who authorised the appeal, when it was received, when the original decision was made, the suspect and the reason for appeal.

The reason is cut to its first 35 words, with a "See more" link that expands it in place and changes to "See less". This keeps the free text the research asked for while keeping the card easy to scan. A "View appeal details" link opens the read-only page.

We first tried an [accordion](https://design-system.service.gov.uk/components/accordion/) holding the full appeal, charging and review details. We removed it because the overview only needs a snapshot. Everyone can already see the full information through Start task or View appeal details.

<!-- SCREENSHOT PLACEHOLDER — duplicate as needed -->
![Screenshot described under image](case-overview-appeal-card.png)

*Caption: The PCD appeal card on the case overview*

Describe the screenshot here. Include the page heading, what each section contains, and what the buttons and links say.
<!-- END PLACEHOLDER -->

### Telling reviewers the case has been reactivated

The appeal page now says plainly: "This case has been reactivated. This task can only be completed by carrying out an appeals case review." The way into case review is a "Continue to case review" button at the bottom of the page, replacing a link that was easy to miss.

We built 2 versions of the message: a [notification banner](https://design-system.service.gov.uk/components/notification-banner/) at the top of the page, and [inset text](https://design-system.service.gov.uk/components/inset-text/) just above the button. Neither has been tested yet.

<!-- SCREENSHOT PLACEHOLDER — duplicate as needed -->
![Screenshot described under image](appeal-reactivated-message.png)

*Caption: The appeal page with the reactivation message and button to case review*

Describe the screenshot here. Include the page heading, what each section contains, and what the buttons and links say.
<!-- END PLACEHOLDER -->

### Keeping reassignment simple

The reassign page has 2 [radios](https://design-system.service.gov.uk/components/radios/): "To an individual" and "To a work assignment team". Participants found this journey clear.

We tried adding "To me" for CPSD's pick-up-from-the-pool model. We dropped it because it was outside the MVP, and because users can already find themselves through "To an individual".

<!-- SCREENSHOT PLACEHOLDER — duplicate as needed -->
![Screenshot described under image](reassign-options.png)

*Caption: The reassign page with 2 options*

Describe the screenshot here. Include the page heading, what each section contains, and what the buttons and links say.
<!-- END PLACEHOLDER -->

## What we found about case reactivation

We had assumed the case was reactivated when a DCP started the task. We then found it is reactivated as soon as the task is created.

That means every appeal reactivates the case, including appeals that won't be upheld. Research participants were already concerned that reactivating a case triggers a new charging decision and counts twice in performance data, even when nothing has changed. The draft to-be process only reactivates a case once an appeal is upheld, so the live behaviour and the process don't yet agree.

## What we have not resolved yet

- **Rejecting at intake.** Charging Managers want to reject out-of-time appeals, or appeals from someone below inspector level, without a full review.
- **Appeals that aren't upheld.** The reactivation message only mentions an appeals case review, but the to-be process has a separate ad hoc review for appeals that aren't upheld.
- **Deadlines.** The research refers to both a 14-day and a 10-day timeframe, and the appeal page doesn't yet show when the original decision was made.
- **Filtering.** The task list filters aren't working yet, and priority appeals look the same as routine ones in the main list.
- **Notifications.** DCPs move between custody suites and courts, so they need notifications and mobile access to build the habit of using tasks.
- **Appeals log and messages to police.** Areas still keep their own spreadsheets and email acknowledgements to police by hand.
- **The 2 reactivation message versions.** We need to test which works better.
