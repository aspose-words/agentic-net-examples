using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and set formatting.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the first line indent to half an inch (36 points).
        builder.ParagraphFormat.FirstLineIndent = 36.0;

        // Write a sample paragraph.
        builder.Writeln("This paragraph has a first line indent of half an inch.");

        // Save the document to the local file system.
        doc.Save("FirstLineIndent.docx");
    }
}
