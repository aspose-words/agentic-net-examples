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

        // Insert a paragraph with some text.
        builder.Writeln("This paragraph will be centered.");

        // Set the alignment of the current paragraph to center.
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Save the document to the file system.
        doc.Save("CenteredParagraph.docx");
    }
}
