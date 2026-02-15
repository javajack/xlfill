import { getCollection } from "astro:content";
import type { CollectionEntry } from "astro:content";

export async function GET() {
  const docs = await getCollection("docs");

  // Build comprehensive documentation content
  let content = `# XLFill - Complete Documentation\n\n`;
  content += `> A Go library for template-first Excel report generation. Design in Excel, fill with Go.\n\n`;
  content += `> This file contains the complete documentation content for LLM consumption.\n\n`;
  content += `---\n\n`;

  // Sort docs by category and then by slug
  const sortedDocs = [...docs].sort((a, b) => {
    const aPath = a.slug || "";
    const bPath = b.slug || "";

    // Index first
    if (aPath === "index") return -1;
    if (bPath === "index") return 1;

    // Then guides
    if (aPath.startsWith("guides/") && !bPath.startsWith("guides/")) return -1;
    if (!aPath.startsWith("guides/") && bPath.startsWith("guides/")) return 1;

    // Then commands
    if (aPath.startsWith("commands/") && !bPath.startsWith("commands/"))
      return -1;
    if (!aPath.startsWith("commands/") && bPath.startsWith("commands/"))
      return 1;

    // Then reference
    if (aPath.startsWith("reference/") && !bPath.startsWith("reference/"))
      return -1;
    if (!aPath.startsWith("reference/") && bPath.startsWith("reference/"))
      return 1;

    return aPath.localeCompare(bPath);
  });

  // Add overview section
  content += `## Overview\n\n`;
  content += `XLFill is a Go library that revolutionizes Excel report generation by using a template-first approach. Instead of writing verbose code to style cells, merge ranges, and format data, you design your report visually in Excel and let XLFill fill it with data.\n\n`;

  content += `### Key Advantages\n\n`;
  content += `- **Visual Design**: Create templates in Excel, Google Sheets, or LibreOffice\n`;
  content += `- **Zero Styling Code**: All formatting is preserved from the template\n`;
  content += `- **Business User Friendly**: Non-developers can update templates\n`;
  content += `- **Formula Support**: Excel formulas are automatically expanded\n`;
  content += `- **Rich Features**: Loops, conditionals, grids, images, merged cells\n`;
  content += `- **Extensible**: Create custom commands for specific needs\n\n`;

  content += `### Installation\n\n`;
  content += `\`\`\`bash\n`;
  content += `go get github.com/javajack/xlfill\n`;
  content += `\`\`\`\n\n`;

  content += `### Basic Usage\n\n`;
  content += `\`\`\`go\n`;
  content += `package main\n\n`;
  content += `import "github.com/javajack/xlfill"\n\n`;
  content += `func main() {\n`;
  content += `    data := map[string]interface{}{\n`;
  content += `        "employees": []Employee{\n`;
  content += `            {Name: "John", Department: "Sales", Salary: 50000},\n`;
  content += `            {Name: "Jane", Department: "Engineering", Salary: 75000},\n`;
  content += `        },\n`;
  content += `    }\n\n`;
  content += `    err := xlfill.Fill("template.xlsx", "output.xlsx", data)\n`;
  content += `    if err != nil {\n`;
  content += `        panic(err)\n`;
  content += `    }\n`;
  content += `}\n`;
  content += `\`\`\`\n\n`;

  content += `---\n\n`;

  // Add all documentation pages
  content += `## Complete Documentation\n\n`;

  for (const doc of sortedDocs) {
    const slug = doc.slug || "unknown";
    content += `### ${doc.data.title}\n\n`;
    if (doc.data.description) {
      content += `**Description**: ${doc.data.description}\n\n`;
    }
    content += `**URL**: https://javajack.github.io/xlfill/${slug}\n\n`;

    // Include the raw body content if available
    if (doc.body) {
      // Clean up the content - remove frontmatter delimiters if present
      let bodyContent = doc.body.trim();
      if (bodyContent.startsWith("---")) {
        // Remove frontmatter
        const parts = bodyContent.split("---");
        if (parts.length >= 3) {
          bodyContent = parts.slice(2).join("---").trim();
        }
      }

      // Limit body content length to avoid extremely large files
      if (bodyContent.length > 5000) {
        bodyContent =
          bodyContent.substring(0, 5000) +
          "\n\n[Content truncated - see full documentation at URL above]";
      }

      content += `${bodyContent}\n\n`;
    }

    content += `---\n\n`;
  }

  // Add API reference summary
  content += `## Core API Functions\n\n`;
  content += `### xlfill.Fill(templatePath, outputPath, data)\n`;
  content += `Fills an Excel template with data and saves the result.\n\n`;
  content += `**Parameters**:\n`;
  content += `- templatePath: Path to the Excel template (.xlsx)\n`;
  content += `- outputPath: Path where the output file will be saved\n`;
  content += `- data: Data structure (map, struct, or any Go value)\n\n`;

  content += `### xlfill.Describe(templatePath)\n`;
  content += `Analyzes a template and returns its structure, including all commands and expressions.\n\n`;

  content += `---\n\n`;

  // Add command reference
  content += `## Template Commands Reference\n\n`;

  content += `### jx:each - Iterate over collections\n`;
  content += `Syntax: jx:each(items=data.employees, direction=RIGHT)\n`;
  content += `Repeats cells for each item in a collection.\n\n`;

  content += `### jx:if - Conditional rendering\n`;
  content += `Syntax: jx:if(test=value > 100, direction=DOWN)\n`;
  content += `Conditionally includes or removes cells based on expressions.\n\n`;

  content += `### jx:grid - 2D data grids\n`;
  content += `Syntax: jx:grid(data=matrix, headers=true)\n`;
  content += `Renders two-dimensional data structures efficiently.\n\n`;

  content += `### jx:image - Insert images\n`;
  content += `Syntax: jx:image(src=data.imageUrl)\n`;
  content += `Inserts images from URLs or file paths.\n\n`;

  content += `### jx:mergeCells - Merge cells dynamically\n`;
  content += `Syntax: jx:mergeCells(cols=2, rows=1)\n`;
  content += `Merges cells based on data-driven logic.\n\n`;

  content += `### jx:updateCell - Modify cell properties\n`;
  content += `Syntax: jx:updateCell(col=2, row=3, value=newValue)\n`;
  content += `Updates cell values or formatting programmatically.\n\n`;

  content += `### jx:autoRowHeight - Auto-adjust row heights\n`;
  content += `Syntax: jx:autoRowHeight\n`;
  content += `Automatically adjusts row heights to fit content.\n\n`;

  content += `---\n\n`;

  // Add footer with project info
  content += `## Project Information\n\n`;
  content += `- **GitHub**: https://github.com/javajack/xlfill\n`;
  content += `- **Documentation**: https://javajack.github.io/xlfill/\n`;
  content += `- **License**: MIT\n`;
  content += `- **Author**: Rakesh Waghela (https://x.com/webiyo, https://www.linkedin.com/in/rakeshwaghela)\n`;
  content += `- **Consultation**: https://topmate.io/rakeshwaghela\n\n`;

  content += `This documentation is optimized for LLM consumption. For the interactive version with examples and visual aids, visit the full documentation site.\n`;

  return new Response(content, {
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "public, max-age=3600",
    },
  });
}
