using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add merge fields for customer name and address.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD for the customer's name.
        builder.InsertField("MERGEFIELD CustomerName");

        // Move to a new line.
        builder.Writeln();

        // Insert a MERGEFIELD for the customer's address.
        builder.InsertField("MERGEFIELD Address");

        // Save the template document to the file system.
        doc.Save("CustomerTemplate.docx");
    }
}
