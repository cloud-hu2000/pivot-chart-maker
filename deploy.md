# Deployment Checklist

This project now includes the baseline SEO files needed for a public tool site:

- `index.html` metadata, Open Graph tags, Twitter card tags, canonical URL, favicon links, manifest, and JSON-LD.
- `public/robots.txt`
- `public/sitemap.xml`
- `public/favicon.svg`
- `public/apple-touch-icon.svg`
- `public/og-image.svg`

## Must Confirm Before Production

The production domain was not provided, so the SEO files currently use this placeholder:

```text
https://www.pivotchartmaker.org/
```

Before deploying, replace that value everywhere if the final domain is different:

- `index.html`: canonical URL, `og:url`, `og:image`, JSON-LD `url`
- `public/robots.txt`: `Sitemap`
- `public/sitemap.xml`: `<loc>`

## Open Graph Image

`public/og-image.svg` is included as a temporary share image. Many crawlers support SVG poorly for OG previews, so create a final 1200 x 630 PNG and update:

- `index.html`: `og:image`
- `index.html`: `twitter:image`

Recommended final path:

```text
/og-image.png
```

## Search Console

After deployment:

- Verify the production domain in Google Search Console.
- Submit `https://www.pivotchartmaker.org/sitemap.xml`.
- Test the homepage with Google's Rich Results Test and URL Inspection.
- Test Open Graph rendering with LinkedIn Post Inspector, Facebook Sharing Debugger, or another card preview tool.

## Coverage Remediation

The sitemap intentionally lists only the six substantial, canonical pages. Earlier
thin chart-generator landing pages are permanently redirected in `vercel.json` to
the closest relevant guide or to the main tool. Do not add those redirected URLs
back to the sitemap or link to them internally.

After the production deployment completes:

- Inspect the sitemap in Search Console and request a recrawl of the homepage.
- Use URL Inspection on one redirected URL and confirm Google sees a single 301
  hop to its intended destination.
- Start "Validate fix" only for `Discovered - currently not indexed`; the
  `Page with redirect` status is expected for legacy URLs and is not an error.

## Contact

The public contact email is:

```text
cloudhu2000@gmail.com
```

Update the email in `src/App.jsx` and `index.html` if this changes.
