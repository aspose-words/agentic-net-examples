using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field at the start of the document.
        builder.InsertParagraph();
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.Writeln(); // Add a blank line after the TOC.

        // Add sample headings and content to generate TOC entries.
        for (int i = 1; i <= 3; i++)
        {
            // Heading (will appear in the TOC).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // Normal paragraph content.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"Content of chapter {i}.");

            // Add extra lines to force pagination.
            for (int j = 0; j < 30; j++)
            {
                builder.Writeln($"Line {j + 1} of chapter {i}.");
            }

            // Insert a page break after each chapter except the last.
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Update all fields in the document (including the TOC).
        doc.UpdateFields();

        // Rebuild page layout so the TOC page numbers are refreshed.
        doc.UpdatePageLayout();

        // Save the resulting document.
        const string outputPath = "RebuiltTOC.docx";
        doc.Save(outputPath);
    }
}
