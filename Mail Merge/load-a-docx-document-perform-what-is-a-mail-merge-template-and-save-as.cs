using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToPng
{
    static void Main()
    {
        // Load the existing DOCX file that contains MERGEFIELD tags.
        Document doc = new Document("Template.docx");

        // Prepare simple mail‑merge data.
        // The field names must match the MERGEFIELD names in the template.
        string[] fieldNames = { "FirstName", "LastName", "Address" };
        object[] fieldValues = { "John", "Doe", "123 Main St., Anytown" };

        // Execute the mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Render the first page of the merged document to a PNG image.
        // Using the overload that specifies the format directly.
        doc.Save("MergedResult.png", SaveFormat.Png);
    }
}
