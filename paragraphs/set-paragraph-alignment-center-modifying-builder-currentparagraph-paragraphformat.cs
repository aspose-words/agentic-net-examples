using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph of text.
        builder.Writeln("This paragraph will be centered.");

        // Set the alignment of the current paragraph to center.
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Save the document to a file.
        doc.Save("CenteredParagraph.docx");
    }
}
