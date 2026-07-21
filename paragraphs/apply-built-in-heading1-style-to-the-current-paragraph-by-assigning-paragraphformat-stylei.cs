using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading1 style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Add some text that will be formatted with Heading1.
        builder.Writeln("This paragraph uses the Heading1 style.");

        // Save the document to the local file system.
        doc.Save("Output.docx");
    }
}
