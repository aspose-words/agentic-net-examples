using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

namespace MailMergeCustomText
{
    // Custom callback that formats the text inserted for each merge field.
    public class CustomFieldMergingCallback : IFieldMergingCallback
    {
        // This method is called for every MERGEFIELD encountered during the mail merge.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Use the field name to decide how to format the inserted text.
            // If the field value is null, treat it as an empty string.
            string value = args.FieldValue?.ToString() ?? string.Empty;

            switch (args.DocumentFieldName)
            {
                case "FirstName":
                    // Example: prepend "First: " to the value.
                    args.Text = $"First: {value}";
                    break;
                case "LastName":
                    // Example: prepend "Last: " to the value.
                    args.Text = $"Last: {value}";
                    break;
                default:
                    // For any other field, insert the raw value.
                    args.Text = value;
                    break;
            }
        }

        // Image merging is not required for this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert merge fields that will be populated from the data source.
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln();

            // Prepare a DataTable that serves as the mail‑merge data source.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("FirstName");
            table.Columns.Add("LastName");
            table.Rows.Add("John", "Doe");
            table.Rows.Add("Jane", "Smith");

            // Assign the custom field‑merging callback.
            doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

            // Execute the mail merge using the DataTable.
            doc.MailMerge.Execute(table);

            // Save the merged document to the file system.
            doc.Save("Merged.docx");
        }
    }
}
