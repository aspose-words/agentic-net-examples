using System;
using Aspose.Words;

class MailMergeExample
{
    static void Main()
    {
        // Load the DOCX template that contains MERGEFIELDs.
        Document doc = new Document("Template.docx");

        // Define the merge field names present in the template and the corresponding values.
        string[] fieldNames = { "FullName", "Address", "City" };
        object[] fieldValues = { "John Doe", "123 Main St.", "Springfield" };

        // Execute a standard (single‑record) mail merge using the flat data source.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the result to a new file.
        doc.Save("MergedOutput.docx");
    }
}
