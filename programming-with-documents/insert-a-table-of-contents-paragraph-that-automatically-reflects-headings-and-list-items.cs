using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsTocExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Attach a DocumentBuilder to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents (TOC) field.
            // \o "1-3"  – include heading levels 1 to 3.
            // \h        – make entries clickable hyperlinks.
            // \z        – hide page numbers in web layout.
            // \u        – use outline levels.
            // \t "List Paragraph,1" – include paragraphs with the "List Paragraph" style (list items).
            string tocSwitches = @"\o ""1-3"" \h \z \u \t ""List Paragraph,1""";
            builder.InsertTableOfContents(tocSwitches);

            // Insert a page break so that the TOC appears on its own page.
            builder.InsertBreak(BreakType.PageBreak);

            // ---------- Add sample headings ----------
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1: Introduction");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1: Overview");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("Subsection 1.1.1: Details");

            // ---------- Add a sample list ----------
            // Create a bulleted list.
            List list = doc.Lists.Add(ListTemplate.BulletDefault);
            builder.ListFormat.List = list;

            builder.Writeln("First list item");
            builder.Writeln("Second list item");
            builder.Writeln("Third list item");

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Update all fields (including the TOC) to reflect the newly added content.
            doc.UpdateFields();

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "DocumentWithToc.docx");
            doc.Save(outputPath);
        }
    }
}
