using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents that includes only "Heading 1" and a custom style "MyCustomStyle".
        // \t switch takes a comma‑separated list of style names.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\t \"Heading 1,MyCustomStyle\"");

        // Add a page break after the TOC so headings start on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Add a Heading 1 paragraph (will appear in the TOC).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        // Add a Heading 2 paragraph (will NOT appear because it's not listed in \t).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Background");

        // Define a custom paragraph style named "MyCustomStyle".
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Size = 16;
        customStyle.Font.Bold = true;

        // Add a paragraph using the custom style (will appear in the TOC).
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("Custom Chapter: Special Topic");

        // Save the document to disk.
        string outputPath = "TOC_Styles.docx";
        doc.Save(outputPath);

        // Verify that the file was saved and can be reopened.
        if (File.Exists(outputPath))
        {
            Document loaded = new Document(outputPath);
            // No further action needed; just confirming successful load.
        }
    }
}
