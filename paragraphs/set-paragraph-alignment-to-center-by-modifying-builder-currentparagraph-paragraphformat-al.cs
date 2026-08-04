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

        // Write a line of text – this creates a paragraph and makes it the current paragraph.
        builder.Writeln("This paragraph will be centered.");

        // Modify the alignment of the current paragraph to center.
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Save the document to a file.
        doc.Save("CenteredParagraph.docx");
    }
}
