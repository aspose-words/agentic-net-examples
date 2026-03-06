using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains MERGEFIELDs.
        Document doc = new Document("Template.docx");

        // Define the names of the merge fields present in the template
        // and the corresponding values to insert.
        string[] fieldNames = { "FullName", "Company", "Address", "City" };
        object[] fieldValues = { "James Bond", "MI5 Headquarters", "Milbank", "London" };

        // Execute a mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as PDF.
        doc.Save("MergedResult.pdf", SaveFormat.Pdf);
    }
}
