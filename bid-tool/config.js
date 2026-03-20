/**
 * ─────────────────────────────────────────────
 *  COMPANY CONFIG — edit this file per client
 * ─────────────────────────────────────────────
 *  To white-label for a new contractor:
 *  1. Update every field below
 *  2. Push to their GitHub Pages repo
 *  That's it.
 * ─────────────────────────────────────────────
 */

const COMPANY_CONFIG = {

  // ── Branding ───────────────────────────────
  companyName:  'Cornerstone',                       // Short name (sidebar, titles)
  companyFull:  'Cornerstone Hardscape & Excavation',// Full legal/display name
  tagline:      'Hardscape & Excavation',            // Shown under sidebar name
  logoUrl:      '',                                  // URL to logo image (leave blank to use text)
  accentColor:  '#0A0A0A',                           // Primary button color

  // ── CRM pipeline stages ────────────────────
  stages: ['New Lead', 'Contacted', 'Quote Sent', 'Follow-Up', 'Won', 'Lost'],

  // ── Lead sources ──────────────────────────
  leadSources: ['Referral', 'Google', 'Facebook', 'Door Hanger', 'Website', 'Nextdoor', 'Other'],

  // ── Bid Tool services ──────────────────────
  // Each service appears as a tab on the bid tool.
  // pricingType: 'daily' — charges a flat day rate divided by acres/day
  //              'per_acre' — charges a flat rate per acre (coming soon)
  //              'fixed' — single fixed price regardless of acreage (coming soon)
  services: [
    {
      id:          'forestry',
      name:        'Forestry Mulching',
      pricingType: 'daily',
      dayRate:     3500,                    // $ per day
      densities: [
        { label: 'Light',  acresPerDay: 2   },
        { label: 'Medium', acresPerDay: 1   },
        { label: 'Dense',  acresPerDay: 0.5 },
      ],
    },
    // ── Add more services below ─────────────
    // {
    //   id:          'grading',
    //   name:        'Land Grading',
    //   pricingType: 'daily',
    //   dayRate:     2800,
    //   densities: [
    //     { label: 'Light',  acresPerDay: 3   },
    //     { label: 'Medium', acresPerDay: 1.5 },
    //     { label: 'Dense',  acresPerDay: 0.75 },
    //   ],
    // },
  ],

};
