using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a label for the date field.
        builder.Writeln("Current date:");

        // Insert a DATE field with the desired format: "MMMM dd, yyyy".
        // The field code is inserted without the surrounding braces.
        builder.InsertField("DATE \\@ \"MMMM dd, yyyy\"");

        // Save the document to the local file system.
        doc.Save("CurrentDate.docx");
    }
}
