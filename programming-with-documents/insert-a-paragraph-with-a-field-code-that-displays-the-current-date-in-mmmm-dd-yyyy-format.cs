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

        // Insert an empty paragraph where the field will be placed.
        builder.InsertParagraph();

        // Insert a DATE field with the desired format: "MMMM dd, yyyy".
        // The field code is provided without the surrounding braces.
        Field dateField = builder.InsertField("DATE \\@ \"MMMM dd, yyyy\"");

        // Update all fields in the document so the DATE field shows the current date.
        doc.UpdateFields();

        // Save the document to the file system.
        doc.Save("DateField.docx");
    }
}
