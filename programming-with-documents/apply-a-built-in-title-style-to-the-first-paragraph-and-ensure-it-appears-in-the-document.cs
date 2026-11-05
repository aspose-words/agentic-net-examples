using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first paragraph (the title).
        builder.Writeln("My Document Title");

        // Apply the built‑in Title style to the first paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;

        // The Title style automatically maps to the outline level for titles,
        // so an explicit OutlineLevel assignment is not required.

        // Add a normal paragraph.
        builder.Writeln("This is the body of the document.");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
