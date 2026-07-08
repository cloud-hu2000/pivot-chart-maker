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
https://pivotchartmaker.com/
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
- Submit `https://YOUR_DOMAIN/sitemap.xml`.
- Test the homepage with Google's Rich Results Test and URL Inspection.
- Test Open Graph rendering with LinkedIn Post Inspector, Facebook Sharing Debugger, or another card preview tool.

## Contact

The public contact email is:

```text
cloudhu2000@gmail.com
```

Update the email in `src/App.jsx` and `index.html` if this changes.
