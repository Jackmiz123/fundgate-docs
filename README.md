# FundGate Document Engine — one codebase, two deployments

This is your normal app. With one environment variable it also runs as the gated
broker workspace. Same code, same contract templates, same disclosure modules.
Update a template once, push, and BOTH deployments pick it up.

## How the mode works
- BROKER_MODE unset  -> your normal app (main form, Past Deals visible, Word + PDF). Unchanged.
- BROKER_MODE=true    -> gated login, PDF-only, Past Deals hidden, dark broker UI,
                         no-state disclosure + 4-tier addendum removed.

Both modes share: contract templates (.docx), disclosure modules, the generation
engine, and the Supabase deals table. So the broker's generated documents are
produced by YOUR app's templates, and his deals write into YOUR Past Deals.

## Deploy: two Render Web Services, same repo
1. Push this folder to your existing GitHub repo (replacing what's there).
   Your current fundgate-weekly service redeploys automatically and behaves
   exactly as before (BROKER_MODE is not set on it).

2. Create the broker service:
   Render -> New + -> Web Service -> connect the SAME repo.
   Under Environment, set:
     BROKER_MODE   = true
     SITE_PASSWORD = Abie1@
     SUPABASE_URL  = (same value as your main service)
     SUPABASE_KEY  = (same value as your main service)
   Deploy. You get a second URL (e.g. fundgate-broker.onrender.com) for the broker.

That's it. From now on, updating a contract template or disclosure = one push =
both sites updated. The broker's docs are always in sync with yours because they
ARE yours.

## Note on adding brand-new form fields
Contract/template/disclosure/clause updates need no form change and sync
automatically. If you ever add a brand-new INPUT field, add it to both
fundgate_form.html (main) and fundgate_form_broker.html (broker) so both can
collect it. Everything else is shared.
