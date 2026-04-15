using System;
using Aspose.Words;

public class ApplyHeadingStyleExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder which simplifies document construction.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading1 style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Write some text; it will be formatted with Heading1 style.
        builder.Writeln("This paragraph uses the Heading1 style.");

        // Save the document to the local file system.
        doc.Save("HeadingStyle.docx");
    }
}
