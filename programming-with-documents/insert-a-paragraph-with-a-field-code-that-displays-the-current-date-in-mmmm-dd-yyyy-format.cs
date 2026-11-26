using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph and add a DATE field with the desired format.
        builder.Writeln(); // Ensure we are in a new paragraph.
        builder.InsertField(@"DATE  \@ ""MMMM dd, yyyy""");

        // Save the document to the local file system.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
