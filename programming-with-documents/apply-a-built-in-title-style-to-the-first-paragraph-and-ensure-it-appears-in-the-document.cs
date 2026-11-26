using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Title style and set an outline level that will appear in the document outline.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1; // Title style does not have a dedicated OutlineLevel enum value.
        builder.Writeln("Document Title");

        // Reset formatting for subsequent paragraphs.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
        builder.Writeln("This is the body of the document.");

        // Save the document to disk.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
