using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a DATE field formatted as "MMMM dd, yyyy".
        builder.Writeln("Current date:");
        builder.InsertField(@"DATE  \@ ""MMMM dd, yyyy""");

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
