using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTocUpdate
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents (TOC) field.
            // The switches configure the TOC to include heading levels 1‑3, add hyperlinks, hide page numbers for hidden entries, etc.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Add some headings that will be captured by the TOC.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");
            builder.Writeln("Section 1.2");

            // Insert a new section (starts on a new page) and add more headings.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 2.1");
            builder.Writeln("Section 2.2");

            // Add a deeper heading level (will not appear in the TOC because of the switches).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;
            builder.Writeln("Subsection 2.2.1");

            // After all modifications, update all fields in the document so the TOC reflects the new headings.
            doc.UpdateFields();

            // Define the output path (in the current working directory).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedToc.docx");

            // Save the document.
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
