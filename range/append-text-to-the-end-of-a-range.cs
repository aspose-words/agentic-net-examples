using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Add a paragraph with initial text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World");

        // Move the builder to the end of the document's range.
        builder.MoveToDocumentEnd();

        // Append the desired text.
        builder.Write(" - Appended text.");

        // Save the document.
        string outputPath = "Result.docx";
        doc.Save(outputPath);
    }
}
