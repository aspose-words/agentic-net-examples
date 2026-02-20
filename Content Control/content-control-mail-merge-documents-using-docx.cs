using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Load a DOCX template that contains content controls (structured document tags)
        // whose tags match the field names we will provide (e.g., "FirstName", "LastName").
        Document doc = new Document("Template.docx");

        // Enable mail merge to work with content controls instead of traditional MERGEFIELD fields.
        doc.MailMerge.UseNonMergeFields = true;

        // Define the field names that correspond to the tags of the content controls.
        string[] fieldNames = { "FirstName", "LastName", "Address" };

        // Provide the values that will be inserted into the matching content controls.
        object[] fieldValues = { "John", "Doe", "123 Main St., Anytown" };

        // Perform the mail merge operation.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Optional: remove any empty paragraphs or unused fields that may remain after merging.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs |
                                      MailMergeCleanupOptions.RemoveUnusedFields;

        // Save the merged document to a new DOCX file.
        doc.Save("MergedDocument.docx");
    }
}
