using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD named "Name".
        builder.InsertField(" MERGEFIELD Name ");

        // Assign a custom callback that will insert the name in bold.
        doc.MailMerge.FieldMergingCallback = new InsertBoldNameCallback();

        // Prepare a simple data source.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name");
        table.Rows.Add("John Doe");
        table.Rows.Add("Jane Smith");

        // Perform the mail merge.
        doc.MailMerge.Execute(table);

        // Save the result.
        doc.Save("BoldNames.docx");
    }

    // Custom callback that writes the merge field value in bold.
    private class InsertBoldNameCallback : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            if (args.DocumentFieldName.Equals("Name", StringComparison.OrdinalIgnoreCase))
            {
                // Move the builder to the location of the merge field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Apply bold formatting and write the field value.
                builder.Font.Bold = true;
                builder.Write(args.FieldValue?.ToString() ?? string.Empty);

                // Prevent the default insertion of the field value.
                args.Text = string.Empty;
            }
        }

        // No image handling required for this example.
        public void ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
