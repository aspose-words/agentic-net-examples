using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("This paragraph will be centered.");

        // Set alignment of the current paragraph to center.
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Save the document.
        doc.Save("CenteredParagraph.docx");
    }
}
