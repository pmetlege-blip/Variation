# Email Parsing Checklist

When Paul forwards an email, work through this checklist **in order**.
Don't skip — missing fields are the source of every later problem.

## A. Identify the project

- [ ] **Which project does this belong to?**
  - Look for: project code, site address, client surname, or job name in
    the subject line, body, or attached invoice.
  - If unclear, ask: *"Which project is this for? (project code or site address)"*
- [ ] **Confirm the project code** from the master project list.

## B. Identify the trigger

Pick one — what caused this variation?

- [ ] **Client request** — client asked for a change or addition.
- [ ] **Site condition** — something was discovered during demolition or
      excavation (e.g. rot, asbestos, undocumented service, structural
      issue).
- [ ] **Council / certifier** — regulatory body required a change.
- [ ] **Architect / engineer / designer** — design change pushed down.
- [ ] **Supplier-driven** — material unavailable, substitute required.
- [ ] **Builder-initiated** — Renovate 8 proposed a better way.
- [ ] **Other** — record what it is.

The trigger is recorded on the register and (in plainer English) on the
variation notice under "Reason for variation."

## C. Identify the scope

- [ ] **One-line description** for the register (e.g. *"Replace rotted floor
      joists under bathroom"*).
- [ ] **Full scope** — what physically gets done, who does it, what's
      included and (importantly) what's NOT included.
- [ ] **Trades involved** (carpenter, plumber, electrician, tiler, etc.).
- [ ] **Materials specified** — brand, size, quantity if relevant.
- [ ] **Drawing reference** if the variation is documented on a marked-up
      plan.

## D. Identify the costs

This is the part most often missing. **Cost to me** is the internal
builder cost. **Price to client** is what they'll be invoiced — markup
applied.

### Cost to me — line items

- [ ] **Materials $** — from invoice or quote.
- [ ] **Labour $** — Renovate 8's own labour (hours × rate).
- [ ] **Subcontractor $** — from sub's quote or invoice.
- [ ] **Plant / hire / disposal $** — skips, scaffold, etc.
- [ ] **Other $** — permits, certificates, engineering reports, etc.

### Markup

- [ ] **Markup %** — set per-variation by Paul. Common range 15–25%.
- [ ] **If Paul hasn't stated a markup**, ask:
      *"What markup do you want applied to VAR-### — cost-to-me is $X?"*

### Price to client (calculated)

- [ ] **Ex-GST price** = Cost-to-me × (1 + markup%).
- [ ] **GST** = ex-GST × 10%.
- [ ] **Inc-GST price** = ex-GST + GST.

## E. Identify the time impact

- [ ] **Days added** to the construction period? (zero is valid — record it
      anyway so the notice is unambiguous).
- [ ] **Days removed** if the variation deletes work?
- [ ] If unknown, ask Paul: *"Does VAR-### affect the program? If so, by
      how many working days?"*

## F. Identify the urgency

- [ ] **Is work already in progress?** If yes, the notice must be issued
      immediately and back-dated to when the work was identified. Note
      the date the variation was identified vs. the date work began.
- [ ] **Was it an emergency** (safety, weather, preventing further damage)?
      Emergencies are an exception under most MBA contracts — the variation
      can start before written acceptance, but must still be papered up
      promptly. Flag this in the notice.

## G. Identify the source documents

- [ ] **Supplier invoices** — PDF, save to `01_Source/`.
- [ ] **Supplier quotes** — PDF, save to `01_Source/`.
- [ ] **Photos** — save to `01_Source/photos/`.
- [ ] **Marked-up plans** — save to `01_Source/`.
- [ ] **The original email** — save as `.eml` to `01_Source/`.

## H. Standard reply to Paul (when info is missing)

If the email is incomplete, use this template:

> Hi Paul,
>
> Picking up the variation for **[PROJECT]** — *[one-line description]*.
> Before I raise the notice I need to confirm:
>
> 1. **Project confirmation** — is this for [PROJECT_CODE]?
> 2. **Cost-to-me** — I see $X for [materials/labour/sub]. Is there
>    anything else (own labour, hire, disposal)?
> 3. **Markup** — what % do you want applied?
> 4. **Time impact** — does this affect the program? If so, by how
>    many working days?
> 5. **Trigger** — was this client-requested, a site condition, or
>    something else?
>
> Once I have these I'll raise VAR-### (Phase 1) and email the notice
> to [client]. Final notice will follow when [pending item] is
> confirmed.
>
> Cheers,

Cut down the question list to what's actually missing — don't ask things
the email already answered.

## I. Pre-flight before issuing the notice

Before generating the Phase 1 notice, confirm:

- [ ] Project code matches the register file path.
- [ ] Variation number is the next in sequence (check the register).
- [ ] Description is plain English, no jargon.
- [ ] Price section is either a confirmed fixed price OR explicitly marked
      `Indicative — to be confirmed`.
- [ ] Time impact is stated (zero is a valid answer).
- [ ] Folder path on the notice matches the actual Dropbox folder.
- [ ] No cost-to-me figures appear anywhere on the notice — only
      price-to-client.

**Double-check the last point.** Cost-to-me leaking onto a client document
is the single most damaging mistake this system can make.
