using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains MERGEFIELDs.
        Document doc = new Document("Template.docx");

        // Define the merge fields present in the template and the values to insert.
        string[] fieldNames = { "FirstName", "LastName", "Address" };
        object[] fieldValues = { "John", "Doe", "123 Main St." };

        // Perform a mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the resulting document as a PNG image (renders the first page).
        doc.Save("MergedOutput.png", SaveFormat.Png);
    }
}
