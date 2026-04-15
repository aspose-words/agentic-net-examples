using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph containing a DATE field formatted as "MMMM dd, yyyy".
        // The field code is inserted without the surrounding braces.
        builder.Writeln(); // start a new paragraph
        builder.InsertField("DATE \\@ \"MMMM dd, yyyy\"");

        // Update all fields to ensure the result is current.
        doc.UpdateFields();

        // Save the document to a file in the current directory.
        string outputPath = "DateField.docx";
        doc.Save(outputPath);
    }
}
