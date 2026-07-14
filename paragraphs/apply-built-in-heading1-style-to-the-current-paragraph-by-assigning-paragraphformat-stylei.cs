using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading1 style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Add text; this paragraph will be formatted as Heading1.
        builder.Writeln("This is a Heading 1 paragraph.");

        // Save the document.
        doc.Save("Heading1Example.docx");
    }
}
