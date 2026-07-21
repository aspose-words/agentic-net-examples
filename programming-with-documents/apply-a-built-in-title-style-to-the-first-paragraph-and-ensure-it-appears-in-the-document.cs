using System;
using System.IO;
using Aspose.Words;

public class ApplyTitleStyle
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in "Title" style to the first paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;

        // Ensure the paragraph appears in the document outline by setting an outline level.
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;

        // Write the title text.
        builder.Writeln("My Document Title");

        // Add a normal paragraph after the title.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
        builder.Writeln("This is the body of the document.");

        // Prepare the output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TitleStyleDocument.docx");
        doc.Save(outputPath);
    }
}
