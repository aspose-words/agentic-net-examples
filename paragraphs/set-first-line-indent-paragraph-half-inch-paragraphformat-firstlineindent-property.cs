using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the first line indent to half an inch.
        // 1 inch = 72 points, so 0.5 inch = 36 points.
        builder.ParagraphFormat.FirstLineIndent = 0.5 * 72; // 36 points

        // Write a paragraph to demonstrate the indent.
        builder.Writeln("This paragraph has a first line indent of half an inch.");

        // Save the document to a file.
        doc.Save("FirstLineIndentHalfInch.docx");
    }
}
