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

        // Prepare a data source with a single column "FullName".
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FullName");
        table.Rows.Add("John Doe");
        table.Rows.Add("Jane Smith");

        // Assign a custom callback that will apply bold formatting to the inserted name.
        doc.MailMerge.FieldMergingCallback = new BoldNameCallback();

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        doc.Save("Result.docx");
    }

    // Callback that formats the merged text in bold.
    private class BoldNameCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Apply bold formatting only to the "FullName" field.
            if (args.DocumentFieldName.Equals("FullName", StringComparison.OrdinalIgnoreCase))
            {
                // Move the builder to the position of the current merge field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Set bold formatting and write the field value.
                builder.Font.Bold = true;
                builder.Write(args.FieldValue?.ToString() ?? string.Empty);

                // Prevent the default insertion of the field value.
                args.Text = string.Empty;
            }
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
