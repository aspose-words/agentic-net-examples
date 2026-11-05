using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will be styled as Heading 1.");

        // Apply the built‑in Heading1 style to the current paragraph.
        Paragraph currentParagraph = builder.CurrentParagraph;
        currentParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
