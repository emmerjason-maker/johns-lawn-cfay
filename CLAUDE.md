# John's Lawn Care — CFAY

Single-file static site for John's lawn mowing service for base housing at
Fleet Activities Yokosuka (CFAY), Japan.

## Stack
- One `index.html` (HTML/CSS/vanilla JS inline). No build step, no framework.
- Hosted on GitHub Pages from `main` / root.
  Live URL: https://emmerjason-maker.github.io/johns-lawn-cfay/
- Custom domain planned later.

## Where to edit
All content lives in the `CONFIG` object near the bottom of `index.html`:
business name, tagline, email, WhatsApp number (digits only, intl format),
Stripe Payment Link, service area, prices, equipment/shed note.
Change content there, not in the markup.

## Current content
- Standard mow — Mow, string trim fence line — $10
- Mow and bag — Clippings raked, bagged, and left in your trash area — $20
- John uses the customer's lawnmower and string trimmer, so he needs access
  to their outdoor shed.
- Contact: WhatsApp (primary) and email.
- Payments: Stripe Payment Link ("customers choose what to pay") with an
  auto-generated QR code; invoices sent from the Stripe dashboard.
  Plan is a separate Stripe account for this business.

## Still to do
- Replace placeholder email and WhatsApp number in CONFIG.
- Create the Stripe Payment Link and paste it into CONFIG.stripePaymentLink.
- Confirm home-based business approval for operating on base.

## Conventions
- Keep it a single self-contained file.
- Never commit Stripe secret keys. The public Payment Link URL is fine.
- When proposing changes, output the complete updated file for copy/paste.
- Design: lawn-stripe hero, Bowlby One + Figtree, greens with a mower-yellow
  pay button. Keep that look.
