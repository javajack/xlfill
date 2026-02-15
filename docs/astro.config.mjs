// @ts-check
import { defineConfig } from "astro/config";
import starlight from "@astrojs/starlight";

// https://astro.build/config
export default defineConfig({
  site: "https://javajack.github.io",
  base: "/xlfill",
  integrations: [
    starlight({
      title: "XLFill",
      customCss: ["./src/styles/custom.css"],
      components: {
        Footer: "./src/components/Footer.astro",
      },
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
  ],
});
