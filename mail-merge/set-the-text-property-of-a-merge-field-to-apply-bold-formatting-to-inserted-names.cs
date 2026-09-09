using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to insert a MERGEFIELD.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertField(" MERGEFIELD Name ");

        // Prepare a data source with names to merge.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name");
        table.Rows.Add("John Doe");
        table.Rows.Add("Jane Smith");

        // Assign a custom callback that formats the merged name in bold.
        doc.MailMerge.FieldMergingCallback = new BoldNameCallback();

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Save the result.
        doc.Save("BoldNames.docx");
    }

    // Callback that writes the field value in bold and suppresses the default insertion.
    private class BoldNameCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Move the builder to the current merge field.
            DocumentBuilder builder = new DocumentBuilder(args.Document);
            builder.MoveToMergeField(args.DocumentFieldName);

            // Apply bold formatting and write the field value.
            builder.Font.Bold = true;
            builder.Write(args.FieldValue?.ToString() ?? string.Empty);

            // Prevent the default insertion of the field value.
            args.Text = string.Empty;
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
