# AdSense readiness remediation

Updated: 2026-08-07

## Current status

The code-level Blocker and High items in the supplied preflight report have been addressed on the `codex/adsense-remediation` branch. The production site is still serving the old build: `/privacy-policy.html`, `/about.html`, `/contact.html`, `/terms-of-use.html`, and the new guide pages currently return 404 online. Merge and deploy this branch before requesting another review.

Google does not publish an 800-word approval rule. The 800–1500 word target below is an internal remediation threshold from the supplied report. The content was expanded with worked examples, verification methods, limitations, and original analysis rather than repeated search phrases.

## Blocker remediation

- **ADS-ELIG-03 / ADS-CONTENT-01 / ADS-CONTENT-02:** Expanded every core tool and guide page to 800+ visible words. Added fixed, traceable sample results, realistic use cases, diagnostic steps, interpretation limits, and three full guide pages. Removed repeated exact-match keyword phrasing from the React long-form sections.
- **ADS-PRIV-01:** Added `public/privacy-policy.html` with local-file processing, hosting logs, cookies, advertising, data use, user choices, children, updates, and contact disclosures.
- **Trust pages:** Added `public/about.html`, `public/contact.html`, and `public/terms-of-use.html`.
- **Navigation:** Added About, Contact, Privacy Policy, and Terms of Use links to the global React footer and to the static HTML of every production page.

## High remediation

- **ADS-CONTENT-03:** All 18 core content/tool pages now contain 807–964 visible words in the production build. Trust pages are intentionally concise where appropriate.
- **ADS-PRIV-02:** The privacy policy now states that third-party vendors, including Google, may use cookies based on prior visits, explains personalized advertising, and links to Google privacy controls.

## Medium remediation

- **ADS-CONTENT-08:** Replaced repetitive exact-match phrases with natural explanations, cases, field-selection logic, validation routines, and limitations.
- **ADS-UX-05:** Trust links are present on all built HTML pages.
- **ADS-TXT-02:** Do not publish a placeholder `ads.txt`. After AdSense provides the publisher ID, create `public/ads.txt` with the exact line supplied in AdSense, normally `google.com, pub-XXXXXXXXXXXXXXXX, DIRECT, f08c47fec0942fa0`, then confirm `/ads.txt` returns HTTP 200.

## Unknown items: owner verification

- [ ] **ADS-ELIG-01:** Confirm the payee is at least 18. If not, apply through a parent or guardian account.
- [ ] **ADS-ELIG-02:** In every Google account you control, check AdSense. Confirm there is only one publisher account; add this site to the existing account instead of registering again.
- [ ] **ADS-OWN-01:** Confirm you can edit `index.html` and the shared deployment source. After adding a temporary test meta tag to `<head>`, deploy and verify it appears in View Source, then remove the test tag.
- [ ] **ADS-OWN-02:** Confirm the domain appears in your registrar account and that you can edit its DNS records. Save a registrar/DNS screenshot for your own records.
- [ ] **ADS-SITE-01:** In AdSense → Sites, add `pivotchartmaker.org`, complete the displayed verification method, and wait until the status no longer says “Requires review” or “Needs attention.”
- [ ] **ADS-SITE-02:** Confirm at least one live verification path works: an AdSense script/meta value in `<head>` or a root-level `ads.txt` file. Open View Source or `/ads.txt` after deployment to verify it is public.
- [ ] **ADS-PROG-01:** Put a written no-self-click rule in the operating checklist. Never click live ads; use Google’s test/preview facilities for layout checks.
- [ ] **ADS-PROG-04:** Review acquisition reports by source/medium. Confirm there is no paid-to-click traffic, traffic exchange, bot traffic, spam email, spam comments, or unexplained referral spike. Retain campaign invoices and UTM records for legitimate paid campaigns.
- [ ] **ADS-PROG-06:** Create an ad placement inventory. Allow ads only on substantial content pages; exclude error pages, empty states, downloads, overlays, email content, and the interactive tool’s transient/empty result states.
- [ ] **ADS-PUB-09:** After `ads.txt` is published, confirm its `pub-` ID exactly matches Account → Settings → Account information and that AdSense shows the site as authorized.
- [ ] **ADS-PRIV-03:** Before enabling ads, inspect URLs, analytics data layers, and ad initialization code. Confirm no email address, phone number, full name, user ID that identifies a person, or spreadsheet cell value is passed in a URL or ad request.
- [ ] **ADS-PRIV-04:** If EEA, UK, or Swiss users can visit, configure a Google-certified CMP in AdSense Privacy & messaging (or another certified CMP) before personalized ads are served. Test accept, reject, and manage-options flows in a fresh browser profile.
- [ ] **ADS-PRIV-05:** Search the code and browser permission prompts for geolocation use. The current code does not request precise location. If that changes, update the policy and obtain consent before collection.
- [ ] **ADS-PRIV-07:** Use the unmodified Google ad tag. Confirm no custom code attempts to write, overwrite, proxy, or inspect cookies on Google-owned domains.
- [ ] **ADS-PRIV-08:** Review audience and remarketing configuration. Confirm no audience is built from health, religion, sexual orientation, or other sensitive information. The current general-purpose tool does not need such audiences.
- [ ] **ADS-PRIV-09:** Confirm the site is not running personalized housing, employment, or credit campaigns targeted by protected personal attributes in the United States or Canada. If no such campaigns exist, record this item as not applicable.
- [ ] **ADS-PRIV-10:** Before personalized advertising is enabled, confirm the privacy policy matches every active advertising/analytics vendor and that the CMP records the required choices. Recheck after adding or changing a vendor.

## Post-deployment recheck

- [ ] Merge the remediation branch into the production branch and deploy it.
- [ ] Confirm `/`, all sitemap URLs, `/robots.txt`, and trust pages return HTTP 200 over HTTPS.
- [ ] Confirm the production homepage footer links to About, Contact, Privacy Policy, and Terms of Use.
- [ ] Confirm every core content page contains structured headings and at least 800 visible words without repetitive phrase lists.
- [ ] Confirm the five originally sampled pages show the new long-form content.
- [ ] Confirm the privacy policy includes Google/third-party cookies, prior-visit advertising, personalization controls, data use, and contact details.
- [ ] Confirm the contact email is monitored and belongs to the site operator.
- [ ] Confirm the sitemap contains all trust and guide pages and submit it again in Search Console.
- [ ] Use Search Console URL Inspection on the homepage, privacy policy, About, Contact, and three representative guides; request indexing after successful live tests.
- [ ] Run the AdSense preflight tool again against production and save the new report.
- [ ] Complete all applicable Unknown checks above and record non-applicable items explicitly.
- [ ] Add AdSense verification only after the new production pages are live.
- [ ] After receiving the real publisher ID, publish and verify `/ads.txt`; never use a placeholder ID.
- [ ] Configure a certified CMP before serving personalized ads to EEA, UK, or Swiss visitors.
- [ ] Keep ads off empty tool states, utility-only screens, error pages, overlays, and thin pages.
- [ ] Do not apply until the live site—not only the local build—passes these checks.
