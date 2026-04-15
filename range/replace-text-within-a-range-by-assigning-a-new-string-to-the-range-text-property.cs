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
        builder.Writeln("Hello World! This is a sample document.");

        // Replace the word "Hello" with "Hi" using the document's range.
        // This demonstrates modifying text within a range.
        doc.Range.Replace("Hello", "Hi");

        // Save the modified document to the local file system.
        const string outputFile = "Output.docx";
        doc.Save(outputFile);
    }
}
