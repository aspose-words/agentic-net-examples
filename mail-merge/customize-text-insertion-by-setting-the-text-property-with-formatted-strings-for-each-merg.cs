using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

namespace MailMergeCustomText
{
    // Custom callback that formats the text inserted for each merge field.
    public class CustomFieldMergingCallback : IFieldMergingCallback
    {
        // Called for each simple merge field.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Example formatting: field name in brackets followed by the value.
            // If the field value is null, insert an empty string.
            string value = args.FieldValue?.ToString() ?? string.Empty;
            args.Text = $"[{args.FieldName}] {value}";
        }

        // Required by the interface, but not used in this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling needed.
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert merge fields that will be populated from the data source.
            builder.InsertField(FieldType.FieldMergeField, true).AsFieldMergeField().FieldName = "FirstName";
            builder.Write(" ");
            builder.InsertField(FieldType.FieldMergeField, true).AsFieldMergeField().FieldName = "LastName";
            builder.Writeln();

            // Set up the custom field merging callback.
            doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

            // Prepare a simple data source.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("FirstName");
            table.Columns.Add("LastName");
            table.Rows.Add("John", "Doe");
            table.Rows.Add("Jane", "Smith");

            // Execute the mail merge using the data table.
            doc.MailMerge.Execute(table);

            // Save the merged document.
            doc.Save("MergedDocument.docx");
        }
    }

    // Extension method to simplify casting to FieldMergeField.
    public static class FieldExtensions
    {
        public static FieldMergeField AsFieldMergeField(this Field field)
        {
            return (FieldMergeField)field;
        }
    }
}
