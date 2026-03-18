using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD that will receive a name.
        builder.InsertField(" MERGEFIELD  Name ");

        // Set up a mail‑merge callback that will insert the name in bold.
        doc.MailMerge.FieldMergingCallback = new BoldNameCallback();

        // Prepare a simple data source with a single name.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name");
        table.Rows.Add("John Doe");

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Save the result.
        doc.Save("BoldNames.docx");
    }

    // Callback that inserts the merge value in bold.
    private class BoldNameCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // The value that would normally be inserted.
            string name = args.FieldValue?.ToString() ?? string.Empty;

            // Move the builder to the position of the current merge field.
            DocumentBuilder builder = new DocumentBuilder(args.Document);
            builder.MoveToMergeField(args.DocumentFieldName);

            // Apply bold formatting and write the name.
            builder.Font.Bold = true;
            builder.Write(name);

            // Prevent the default insertion of the field value.
            args.Text = string.Empty;
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
