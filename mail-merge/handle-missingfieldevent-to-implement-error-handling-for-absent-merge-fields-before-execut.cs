using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two merge fields: FirstName and LastName.
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");

        // Prepare a data source that lacks the LastName column.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FirstName");
        table.Rows.Add("John");
        table.Rows.Add("Jane");

        // Subscribe to the FieldMergingCallback to handle absent merge fields.
        doc.MailMerge.FieldMergingCallback = new MissingFieldHandler();

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Save the merged document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.docx");
        doc.Save(outputPath);
    }

    // Implements IFieldMergingCallback to provide custom handling for missing fields.
    private class MissingFieldHandler : IFieldMergingCallback
    {
        // Called for each merge field encountered during mail merge.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // If the data source does not contain a value for this field, FieldValue will be null.
            if (args.FieldValue == null)
            {
                // Log the missing field name (optional).
                Console.WriteLine($"Missing field encountered: {args.FieldName}");

                // Insert placeholder text for the missing field.
                args.Text = $"[Missing {args.FieldName}]";
            }
        }

        // No special handling for image fields in this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Intentionally left blank.
        }
    }
}
