using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some text.
        builder.Writeln("Hello World!");

        // Replace the word "World" with "Aspose" using the document's range.
        doc.Range.Replace("World", "Aspose");

        // Save the modified document to the local file system.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Write the resulting text to the console (no user interaction required).
        Console.WriteLine("Modified document text: " + doc.GetText().Trim());
    }
}
