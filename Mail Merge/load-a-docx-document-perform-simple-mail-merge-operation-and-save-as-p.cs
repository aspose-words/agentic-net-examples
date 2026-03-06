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
        string[] fieldNames = { "FullName", "Address" };
        object[] fieldValues = { "John Doe", "123 Main St, Anytown" };

        // Perform a simple mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as a PNG image.
        // This will render the first page of the document to the image.
        doc.Save("MergedOutput.png", SaveFormat.Png);
    }
}
