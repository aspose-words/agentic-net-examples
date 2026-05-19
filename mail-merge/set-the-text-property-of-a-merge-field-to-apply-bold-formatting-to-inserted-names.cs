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

        // Insert a MERGEFIELD named "Name".
        builder.InsertField(" MERGEFIELD Name ");

        // Prepare a simple data source with one record.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name");
        table.Rows.Add("John Doe");

        // Assign a custom callback that will insert the name in bold.
        doc.MailMerge.FieldMergingCallback = new BoldNameCallback();

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        doc.Save("Output.docx");
    }

    // Custom callback that inserts the merge field value in bold.
    private class BoldNameCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Only handle the "Name" field.
            if (args.DocumentFieldName.Equals("Name", StringComparison.OrdinalIgnoreCase))
            {
                // Move the builder to the position of the merge field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Apply bold formatting and write the field value.
                builder.Font.Bold = true;
                builder.Write(args.FieldValue?.ToString() ?? string.Empty);

                // Prevent the default insertion by clearing the Text property.
                args.Text = string.Empty;
            }
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
