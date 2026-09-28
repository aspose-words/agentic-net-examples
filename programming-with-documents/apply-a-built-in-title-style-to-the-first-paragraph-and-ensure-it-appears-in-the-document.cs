using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in 'Title' style to the first paragraph and ensure it appears in the outline.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
        builder.Writeln("Document Title");

        // Reset style to normal for subsequent paragraphs.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
        builder.Writeln("This is the body of the document.");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
