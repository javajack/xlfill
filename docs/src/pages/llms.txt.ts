import { getCollection } from "astro:content";
import type { CollectionEntry } from "astro:content";

interface GroupedDocs {
  [category: string]: CollectionEntry<"docs">[];
}

export async function GET() {
  const docs = await getCollection("docs");

  // Build llms.txt content
  let content = `# XLFill Documentation\n\n`;
  content += `> A Go library for template-first Excel report generation. Design in Excel, fill with Go.\n\n`;

  // Group docs by category based on their path structure
  const grouped: GroupedDocs = {
    Introduction: [],
    "Template Guide": [],
    Commands: [],
    Advanced: [],
    Reference: [],
  };

  docs.forEach((doc) => {
    const slug = doc.slug || "";
    if (slug === "index") {
      // Skip the index page
      return;
    } else if (
      slug.startsWith("guides/why-xlfill") ||
      slug.startsWith("guides/getting-started") ||
      slug.startsWith("guides/how-templates-work")
    ) {
      grouped["Introduction"].push(doc);
    } else if (
      slug.startsWith("guides/expressions") ||
      slug.startsWith("guides/commands-overview")
    ) {
      grouped["Template Guide"].push(doc);
    } else if (slug.startsWith("commands/")) {
      grouped["Commands"].push(doc);
    } else if (slug.startsWith("guides/")) {
      grouped["Advanced"].push(doc);
    } else if (slug.startsWith("reference/")) {
      grouped["Reference"].push(doc);
    }
  });

  // Generate content for each category
  for (const [category, pages] of Object.entries(grouped)) {
    if (pages.length === 0) continue;

    content += `## ${category}\n`;
    for (const page of pages) {
      const title = page.data.title;
      const description = page.data.description || "Documentation page";
      const url = `https://javajack.github.io/xlfill/${page.slug}`;
      content += `- [${title}](${url}): ${description}\n`;
    }
    content += "\n";
  }

  // Add key features and use cases
  content += `## Key Features\n`;
  content += `- Template-first approach: Design in Excel, fill with Go\n`;
  content += `- Preserves all Excel formatting (fonts, colors, borders, merged cells)\n`;
  content += `- Support for loops (jx:each), conditionals (jx:if), and grids (jx:grid)\n`;
  content += `- Formula support with automatic expansion\n`;
  content += `- Image insertion with automatic sizing\n`;
  content += `- Custom command extensibility\n`;
  content += `- Zero styling code required\n\n`;

  content += `## Installation\n`;
  content += `\`\`\`bash\n`;
  content += `go get github.com/javajack/xlfill\n`;
  content += `\`\`\`\n\n`;

  content += `## Quick Start\n`;
  content += `1. Design your Excel template with expressions like \${data.Field}\n`;
  content += `2. Add commands in cell comments (e.g., jx:each for loops)\n`;
  content += `3. Call xlfill.Fill(templatePath, outputPath, data) from Go\n`;
  content += `4. Get a fully formatted Excel file with your data\n\n`;

  content += `## Project Links\n`;
  content += `- GitHub: https://github.com/javajack/xlfill\n`;
  content += `- Documentation: https://javajack.github.io/xlfill/\n`;
  content += `- Author: Rakesh Waghela (https://x.com/webiyo)\n`;

  return new Response(content, {
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "public, max-age=3600",
    },
  });
}
