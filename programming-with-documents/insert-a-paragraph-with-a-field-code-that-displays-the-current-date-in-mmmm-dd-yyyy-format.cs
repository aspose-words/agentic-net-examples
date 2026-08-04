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

        // Start a new paragraph.
        builder.InsertParagraph();

        // Insert a DATE field with the desired format: "MMMM dd, yyyy".
        // The field code is inserted without the surrounding braces.
        builder.InsertField("DATE \\@ \"MMMM dd, yyyy\"");

        // End the paragraph (adds a paragraph break after the field).
        builder.Writeln();

        // Ensure the field result is up‑to‑date.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("DateField.docx");
    }
}
