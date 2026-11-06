using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph has a first line indent of half an inch.");

        // Retrieve the paragraph that was just added.
        Paragraph paragraph = builder.CurrentParagraph;

        // Set first line indent to half an inch (1 inch = 72 points).
        paragraph.ParagraphFormat.FirstLineIndent = 36; // 0.5 inch = 36 points

        // Save the document to a file.
        doc.Save("FirstLineIndent.docx");
    }
}
