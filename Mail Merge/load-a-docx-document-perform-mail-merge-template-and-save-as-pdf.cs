using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX template.
        Document doc = new Document("Template.docx");

        // Define the merge fields present in the template and the values to insert.
        string[] fieldNames = { "FullName", "Address", "City" };
        object[] fieldValues = { "John Doe", "123 Main St.", "Metropolis" };

        // Perform a mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as PDF.
        doc.Save("Result.pdf", SaveFormat.Pdf);
    }
}
