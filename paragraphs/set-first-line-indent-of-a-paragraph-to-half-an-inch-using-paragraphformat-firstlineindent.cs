using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Insert a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph has a first line indent of half an inch.");

        // Retrieve the inserted paragraph (the first paragraph in the document).
        Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

        // Set the first line indent to half an inch.
        // 1 inch = 72 points, so half an inch = 36 points.
        paragraph.ParagraphFormat.FirstLineIndent = 36f;

        // Save the document to a file.
        doc.Save("FirstLineIndent.docx");
    }
}
