using System;
using Aspose.Words;

namespace ParagraphOutlineLevelExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents that will pick up headings and outline levels.
            // The switches configure the TOC to include levels 1‑3, add hyperlinks, hide page numbers for hidden entries, etc.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // Add a page break so the TOC appears on its own page.
            builder.InsertBreak(BreakType.PageBreak);

            // Insert a main heading using the built‑in Heading 1 style (outline level 1).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Main Heading");

            // Insert a subheading. Instead of using a Heading style, set the outline level directly to Level2.
            // Reset the style to Normal to avoid the built‑in heading style.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
            builder.Writeln("Subheading with Outline Level 2");

            // Add a regular paragraph after the subheading.
            builder.ParagraphFormat.ClearFormatting(); // Reset formatting for normal text.
            builder.Writeln("This is a regular paragraph following the subheading.");

            // Update all fields (including the TOC) so that the table of contents reflects the inserted headings.
            doc.UpdateFields();

            // Save the document to the local file system.
            doc.Save("ParagraphOutlineLevel.docx");
        }
    }
}
