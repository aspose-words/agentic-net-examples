using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph with a DATE field formatted as "MMMM dd, yyyy".
        // Field code syntax: DATE  \@ "MMMM dd, yyyy"
        builder.Writeln(); // start a new paragraph
        builder.InsertField(@"DATE  \@ ""MMMM dd, yyyy""");

        // Update fields so the DATE field shows the current date.
        doc.UpdateFields();

        // Save the document to a file.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
