// @ts-check
import { defineConfig } from "astro/config";
import starlight from "@astrojs/starlight";
import sitemap from "@astrojs/sitemap";

// https://astro.build/config
export default defineConfig({
  site: "https://javajack.github.io",
  base: "/xlfill",
  integrations: [
    starlight({
      title: "XLFill",
      description:
        "A Go library for template-first Excel report generation. Design in Excel, fill with Go.",
      customCss: ["./src/styles/custom.css"],
      components: {
        Footer: "./src/components/Footer.astro",
        Head: "./src/components/Head.astro",
      },
      head: [
        // Google Consent Mode v2 - MUST load BEFORE gtag.js (synchronous)
        {
          tag: "script",
          content: `
            // Initialize dataLayer and gtag function
            window.dataLayer = window.dataLayer || [];
            function gtag(){dataLayer.push(arguments);}

            // Regional scoping: Detect if user is in GDPR region
            function isGDPRRegion() {
              const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
              const euTimezones = ['Europe/', 'Atlantic/Reykjavik', 'Atlantic/Azores', 'Atlantic/Madeira'];
              return euTimezones.some(zone => tz.startsWith(zone));
            }

            // Set consent defaults based on region
            const isGDPR = isGDPRRegion();

            gtag('consent', 'default', {
              'ad_storage': 'denied',
              'ad_user_data': 'denied',
              'ad_personalization': 'denied',
              'analytics_storage': isGDPR ? 'denied' : 'granted',
              'functionality_storage': 'granted',
              'personalization_storage': 'denied',
              'security_storage': 'granted',
              'wait_for_update': 500,
            });

            // Store GDPR flag for banner logic
            window.__isGDPRRegion = isGDPR;
          `,
        },
        // Google Analytics - Load gtag.js (async, after consent default)
        {
          tag: "script",
          attrs: { async: true, src: "https://www.googletagmanager.com/gtag/js?id=G-G986QLPFZ1" },
        },
        {
          tag: "script",
          content: `
            window.dataLayer = window.dataLayer || [];
            function gtag(){dataLayer.push(arguments);}
            gtag('js', new Date());
            gtag('config', 'G-G986QLPFZ1', {
              'anonymize_ip': true,
              'cookie_flags': 'SameSite=None;Secure'
            });
          `,
        },
        // Default Open Graph image for all pages
        {
          tag: "meta",
          attrs: {
            property: "og:image",
            content: "https://javajack.github.io/xlfill/og-image.svg",
          },
        },
        {
          tag: "meta",
          attrs: {
            property: "og:image:width",
            content: "1200",
          },
        },
        {
          tag: "meta",
          attrs: {
            property: "og:image:height",
            content: "630",
          },
        },
        {
          tag: "meta",
          attrs: {
            property: "og:image:type",
            content: "image/svg+xml",
          },
        },
        {
          tag: "meta",
          attrs: {
            name: "twitter:image",
            content: "https://javajack.github.io/xlfill/og-image.svg",
          },
        },
        {
          tag: "meta",
          attrs: {
            name: "twitter:card",
            content: "summary_large_image",
          },
        },
        {
          tag: "meta",
          attrs: {
            name: "twitter:site",
            content: "@webiyo",
          },
        },
        // Additional SEO meta tags
        {
          tag: "meta",
          attrs: {
            name: "keywords",
            content:
              "excel, go, golang, template, report generation, xlsx, spreadsheet, data export, excel library, excel automation",
          },
        },
        {
          tag: "meta",
          attrs: {
            name: "author",
            content: "Rakesh Waghela",
          },
        },
        // Yandex Webmaster verification
        { tag: "meta", attrs: { name: "yandex-verification", content: "5281e40eca9463d2" } },
        // Cloudflare Web Analytics
        { tag: "script", attrs: { defer: true, src: "https://static.cloudflareinsights.com/beacon.min.js", "data-cf-beacon": '{"token": "7ce325bb227e4b42a8406f369ff4e788"}' } },
        // Cookie Consent Banner (injected via script)
        { tag: "script", attrs: { defer: true, src: "/xlfill/cookie-consent.js" } },
        // Structured Data (JSON-LD)
        {
          tag: "script",
          attrs: { type: "application/ld+json" },
          content: JSON.stringify({
            "@context": "https://schema.org",
            "@graph": [
              {
                "@type": "WebSite",
                name: "XLFill",
                url: "https://javajack.github.io/xlfill/",
                description: "A Go library for template-first Excel report generation. Design in Excel, fill with Go.",
                author: { "@id": "#rakesh" },
              },
              {
                "@type": "Person",
                "@id": "#rakesh",
                name: "Rakesh Waghela",
                url: "https://www.linkedin.com/in/rakeshwaghela",
                sameAs: [
                  "https://x.com/webiyo",
                  "https://www.linkedin.com/in/rakeshwaghela",
                  "https://topmate.io/rakeshwaghela",
                ],
              },
            ],
          }),
        },
      ],
      social: [
        {
          icon: "github",
          label: "GitHub",
          href: "https://github.com/javajack/xlfill",
        },
        {
          icon: "x.com",
          label: "Rakesh on X",
          href: "https://x.com/webiyo",
        },
        {
          icon: "linkedin",
          label: "Rakesh on LinkedIn",
          href: "https://www.linkedin.com/in/rakeshwaghela",
        },
        {
          icon: "external",
          label: "Book a Consultation",
          href: "https://topmate.io/rakeshwaghela",
        },
      ],
      sidebar: [
        {
          label: "Introduction",
          items: [
            { label: "Why XLFill?", slug: "guides/why-xlfill" },
            { label: "Getting Started", slug: "guides/getting-started" },
            { label: "How Templates Work", slug: "guides/how-templates-work" },
          ],
        },
        {
          label: "Template Guide",
          items: [
            { label: "Expressions", slug: "guides/expressions" },
            { label: "Commands Overview", slug: "guides/commands-overview" },
          ],
        },
        {
          label: "Commands",
          items: [
            { label: "jx:area", slug: "commands/area" },
            { label: "jx:each", slug: "commands/each" },
            { label: "jx:if", slug: "commands/if" },
            { label: "jx:grid", slug: "commands/grid" },
            { label: "jx:image", slug: "commands/image" },
            { label: "jx:mergeCells", slug: "commands/mergecells" },
            { label: "jx:updateCell", slug: "commands/updatecell" },
            { label: "jx:autoRowHeight", slug: "commands/autorowheight" },
          ],
        },
        {
          label: "Advanced",
          items: [
            { label: "Formulas", slug: "guides/formulas" },
            { label: "Custom Commands", slug: "guides/custom-commands" },
            { label: "Area Listeners", slug: "guides/area-listeners" },
            {
              label: "Debugging & Troubleshooting",
              slug: "guides/debugging",
            },
          ],
        },
        {
          label: "Reference",
          items: [
            { label: "Examples", slug: "reference/examples" },
            { label: "API Reference", slug: "reference/api" },
            { label: "Performance", slug: "reference/performance" },
          ],
        },
      ],
    }),
    sitemap({
      // Filter out any search or dynamic pages
      filter: (page) => !page.includes("/search") && !page.includes("?"),
      // Customize URLs if needed
      customPages: [
        "https://javajack.github.io/xlfill/llms.txt",
        "https://javajack.github.io/xlfill/llms-full.txt",
      ],
    }),
  ],
});
