using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the first line indent to half an inch (36 points).
        builder.ParagraphFormat.FirstLineIndent = 36;

        // Add a paragraph to demonstrate the indent.
        builder.Writeln("This paragraph has a first line indent of half an inch.");

        // Save the document to the current directory.
        doc.Save("FirstLineIndentHalfInch.docx");
    }
}
