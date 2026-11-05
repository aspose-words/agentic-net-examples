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

        // Insert a Table of Contents field at the beginning of the document.
        // The TOC will include headings of levels 1 through 3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Insert a page break so that the following content starts on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Add several sections with headings that will be captured by the TOC.
        for (int i = 1; i <= 3; i++)
        {
            // Set the style to Heading 1 for the chapter title.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // Add some body text under the heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}. It contains some sample text to demonstrate the Table of Contents update functionality.");
        }

        // After adding new headings, update all fields in the document (including the TOC).
        doc.UpdateFields();

        // Save the updated document to disk.
        string outputPath = "UpdatedTOC.docx";
        doc.Save(outputPath);
    }
}
