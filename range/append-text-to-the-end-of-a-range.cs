using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add initial content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original text.");

        // Append additional text to the end of the document's range.
        DocumentBuilder appendBuilder = new DocumentBuilder(doc);
        appendBuilder.MoveToDocumentEnd();
        appendBuilder.Writeln("Appended text.");

        // Save the resulting document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Output the combined text to the console for verification.
        Console.WriteLine("Document text after appending:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
