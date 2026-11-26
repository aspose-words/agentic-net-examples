using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("This is a sample paragraph with a first line indent.");

        // Retrieve the paragraph that was just added.
        Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

        // Set the first line indent to half an inch (0.5 * 72 points = 36 points).
        paragraph.ParagraphFormat.FirstLineIndent = 36.0;

        // Save the document to a file.
        doc.Save("FirstLineIndent.docx");
    }
}
