using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class MergeFieldsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert MERGEFIELD fields into the template.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a merge field for the first name.
        builder.InsertField(" MERGEFIELD FirstName ");

        // Insert a space between the fields.
        builder.Write(" ");

        // Insert a merge field for the last name.
        builder.InsertField(" MERGEFIELD LastName ");

        // Add a paragraph break after the fields.
        builder.InsertParagraph();

        // Prepare a data source – a DataTable with columns matching the merge field names.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");

        // Add sample rows.
        table.Rows.Add("John", "Doe");
        table.Rows.Add("Jane", "Smith");

        // Execute the mail merge using the DataTable as the data source.
        doc.MailMerge.Execute(table);

        // Save the merged document to disk.
        doc.Save("MergedDocument.docx");
    }
}
