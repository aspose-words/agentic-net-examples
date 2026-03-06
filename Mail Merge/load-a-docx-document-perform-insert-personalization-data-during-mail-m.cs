using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Template.docx");

        // Define the mail‑merge field names present in the document
        // and the corresponding values to insert.
        string[] fieldNames = { "FirstName", "LastName", "Address" };
        object[] fieldValues = { "John", "Doe", "123 Main St" };

        // Execute a mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Render the first page of the merged document to a PNG image.
        doc.Save("Result.png", SaveFormat.Png);
    }
}
