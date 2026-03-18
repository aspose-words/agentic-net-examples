using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the first paragraph – this will become the title.
        builder.Writeln("My Document Title");

        // Retrieve the first paragraph that was just created.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // Apply the built‑in Title style.
        firstParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;

        // Make the paragraph appear in the document outline by setting its outline level.
        firstParagraph.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;

        // Add additional content so the document has more than just the title.
        builder.Writeln("This is the body of the document.");

        // Save the document to disk.
        doc.Save("TitleOutline.docx");
    }
}
