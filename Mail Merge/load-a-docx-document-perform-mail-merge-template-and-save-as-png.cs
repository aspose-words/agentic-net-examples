using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains MERGEFIELDs.
        Document doc = new Document("Template.docx");

        // Define the merge field names present in the template and the values to insert.
        string[] fieldNames = { "FullName", "Address", "City" };
        object[] fieldValues = { "John Doe", "123 Main St.", "Metropolis" };

        // Perform a mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the resulting document as a PNG image.
        // The Save method with SaveFormat.Png renders the first page of the document to an image.
        doc.Save("MergedDocument.png", SaveFormat.Png);
    }
}
