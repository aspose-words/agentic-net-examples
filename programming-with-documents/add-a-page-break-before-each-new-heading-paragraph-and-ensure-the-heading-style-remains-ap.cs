using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper method to add a heading with a page break before it.
        void AddHeading(string text, StyleIdentifier styleId)
        {
            // Insert an explicit page break before the heading.
            builder.InsertBreak(BreakType.PageBreak);

            // Apply the desired heading style.
            builder.ParagraphFormat.StyleIdentifier = styleId;

            // Write the heading text and finish the paragraph.
            builder.Writeln(text);
        }

        // Add some content before the first heading.
        builder.Writeln("Introduction paragraph without a heading.");

        // Add headings; each will start on a new page.
        AddHeading("Chapter 1: Getting Started", StyleIdentifier.Heading1);
        builder.Writeln("Content of chapter 1.");

        AddHeading("Section 1.1: Overview", StyleIdentifier.Heading2);
        builder.Writeln("Details for section 1.1.");

        AddHeading("Chapter 2: Advanced Topics", StyleIdentifier.Heading1);
        builder.Writeln("Content of chapter 2.");

        // Ensure the builder's paragraph format is reset to normal for any following text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Define output path and make sure the directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
