using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text to the document.
        builder.Writeln("Original content.");

        // Move the builder cursor to the very start of the document.
        builder.MoveToDocumentStart();

        // Insert new text at the beginning of the document's range.
        builder.Write("Inserted at start. ");

        // Save the resulting document.
        const string outputFile = "Output.docx";
        doc.Save(outputFile);

        // Print the final document text to the console for verification.
        Console.WriteLine(doc.GetText().Trim());
    }
}
