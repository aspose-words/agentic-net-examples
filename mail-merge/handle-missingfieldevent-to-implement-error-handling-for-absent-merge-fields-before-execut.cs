using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeMissingFieldDemo
{
    // Custom callback to handle missing merge fields.
    public class MissingFieldHandler : IFieldMergingCallback
    {
        // This method is called for each merge field encountered during the mail merge.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // If the data source does not contain a value for this field, FieldValue will be null.
            if (args.FieldValue == null)
            {
                // Log the missing field name.
                Console.WriteLine($"Missing merge field detected: {args.DocumentFieldName}");

                // Provide a default value so the merge can continue without errors.
                args.Text = "N/A";
            }
        }

        // No image handling required for this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a MERGEFIELD that will be missing from the data source.
            builder.InsertField("MERGEFIELD FirstName");

            // Prepare a data source that does NOT contain the "FirstName" column.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("LastName");
            table.Rows.Add("Doe");

            // Assign the custom callback to handle missing fields.
            doc.MailMerge.FieldMergingCallback = new MissingFieldHandler();

            // Perform the mail merge. The callback will supply a default value for "FirstName".
            doc.MailMerge.Execute(table);

            // Save the resulting document.
            doc.Save("Merged.docx");
        }
    }
}
