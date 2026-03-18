using System;
using Aspose.Words;

class ApplyHeadingStyle
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which simplifies building the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading1 style to the paragraph that will be written next.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Write the paragraph text. The style applied above will be used for this paragraph.
        builder.Writeln("This paragraph is formatted as Heading 1.");

        // Save the document to a file (adjust the path as needed).
        doc.Save("Heading1.docx");
    }
}
