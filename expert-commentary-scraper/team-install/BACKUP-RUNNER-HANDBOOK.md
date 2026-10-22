# Setting up as a backup runner

This is for whoever installs the Expert Commentary Scraper as a backup, so the shared workbook
keeps updating on days Michael is out (PTO, sick, travel, etc.). Read this before you run the
installer in this folder.

## The short version

- **Website sources: fully automatic.** Nothing to set up. Every installed copy of this tool
  checks the exact same 8 websites (plus 7 event pages) — you'll capture those regardless of your
  own email subscriptions.
- **Email vendor sources: need action from you.** This tool only reads emails that land in **your
  own** Outlook inbox — it doesn't pull from a shared mailbox or from Michael's inbox. It doesn't
  matter who a specific newsletter is addressed to, or what its subject line says — it just has to
  be sitting in your inbox. The lists below tell you exactly which newsletters/distribution lists
  you need to be signed up for so nothing is missed.

## Email vendors — make sure you're subscribed to all of these

### Signum Global Advisors (@signumglobal.com)
- The regular country/topic research notes (subject starts "SIGNUM: ...") — this is their main
  output, make sure you're on the general distribution list.
- **Signum Weekly Digest**
- **Week Ahead**
- Event invitations: **SIGNUM Zoom**, **SIGNUM Investor Trip**, and city-based **Policy Day**
  invites (e.g. "SIGNUM London Policy Day") — worth being on these too, though lower priority than
  the research notes above.

### Eurasia Group (@eurasiagroup.net)
- **EG Daily Brief** — the main recurring roundup, highest priority.
- **EG Update**
- **EG Political Risk Monitor** (monthly, subject usually "Political Risk Data – \<Month\> \<Year\>")
- **EG Global Macro**
- **EG Geo-technology**
- **EG Conference Call** invitations and **EG Conference Call Notes** (the follow-up summary after
  a call happens)
- Eurasia Group's general research notes, tagged per country/practice area in the subject line
  (e.g. "EG China: ...", "EG Germany: ...") — these aren't separate newsletters, they're all part
  of being on Eurasia Group's one general distribution list, which is what you actually need to
  confirm you're on.

### Middle East Council on Global Affairs (@mecouncil.org)
- **Daily Brief** — their main recurring regional roundup, highest priority.
- Any other standalone piece they send — subscribe to their general distribution list.

### Vriens & Partners (@vrienspartners.com)
- No specific named newsletter is tracked yet (no sample has been seen from them). Just make sure
  you're on their general distribution list — every email from this domain is picked up.

## Complementary sources (web + email)

These organizations are already fully covered by the automatic website scraping — their regular
articles need no subscription at all. But they also separately email **event invitations**
(briefings, webinars, calls) that only get captured if you personally receive them, so subscribe to
these too for full coverage:

- **CSIS** (@csis.org) — event/webinar invitations
- **Atlantic Council** (@atlanticcouncil.org) — event invitations (e.g. Global Citizen Awards)
- **CFR** (@cfr.org) — event invitations
- **Lowy Institute** (@lowyinstitute.org) — event invitations

There's no single named newsletter to ask for here — just make sure you're generally subscribed to
each organization's mailing list so their event invites land in your inbox alongside everything
else.

## Before you install

1. Check with Michael (mivogel@deloitte.com) on how to get added to each newsletter/list above, if
   you aren't already receiving it.
2. Confirm at least one real email from each has actually landed in your Outlook inbox —
   subscriptions can take a few days to activate.
3. Then follow the regular install steps in `TEAM-INSTALL-README.md` in this same folder.

## How to tell if it's working

After your first run, check the **Run Log** sheet in the shared workbook — it lists a per-source
count for every run, including each vendor. If a vendor consistently shows 0 while you know they
sent something that day, that's the sign to go check your subscription, not the tool.
