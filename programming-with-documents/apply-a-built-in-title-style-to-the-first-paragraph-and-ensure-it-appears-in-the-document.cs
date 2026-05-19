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

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in "Title" style to the first paragraph.
        builder.ParagraphFormat.StyleName = "Title";

        // Ensure the paragraph appears in the outline by setting its outline level.
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;

        // Write the title text.
        builder.Writeln("Document Title");

        // Add a normal paragraph after the title.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
        builder.Writeln("This is the body of the document.");

        // Determine an output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TitleStyle.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
