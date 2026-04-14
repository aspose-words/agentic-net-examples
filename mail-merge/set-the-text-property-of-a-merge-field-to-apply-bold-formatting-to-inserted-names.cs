using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD that will receive a name.
        builder.InsertField(" MERGEFIELD FullName ");

        // Prepare a simple data source with a single column "FullName".
        DataTable table = new DataTable("Names");
        table.Columns.Add("FullName");
        table.Rows.Add("John Doe");
        table.Rows.Add("Jane Smith");

        // Attach a callback that will write the merged name in bold.
        doc.MailMerge.FieldMergingCallback = new BoldNameCallback();

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Save the result.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "BoldNames.docx");
        doc.Save(outputPath);
    }

    // Callback that inserts the field value in bold.
    private class BoldNameCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Move the builder to the location of the current merge field.
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
