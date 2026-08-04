using System;
using System.IO;
using Aspose.Words;

public class ApplyHeadingStyle
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading1 style to the paragraph that will be created.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Write the heading text; this creates a paragraph with the Heading1 style.
        builder.Writeln("Heading 1 Example");

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Heading1Example.docx");
        doc.Save(outputPath);
    }
}
