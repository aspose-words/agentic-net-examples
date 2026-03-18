using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeWordsMailMergeCustomText
{
    // Custom callback that formats the text inserted for each merge field.
    public class CustomFieldMergingCallback : IFieldMergingCallback
    {
        // This method is called for every simple MERGEFIELD encountered during mail merge.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Example formatting: "[FIELDNAME]: value"
            // args.FieldName – name of the field in the data source.
            // args.FieldValue – original value from the data source.
            // Setting args.Text overrides the default insertion behavior.
            args.Text = $"[{args.FieldName.ToUpper()}]: {args.FieldValue}";
        }

        // No special handling for image fields in this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert three merge fields that will be populated from the data source.
            builder.InsertField("MERGEFIELD FirstName");
            builder.Writeln();
            builder.InsertField("MERGEFIELD LastName");
            builder.Writeln();
            builder.InsertField("MERGEFIELD Salary");
            builder.Writeln();

            // Assign the custom callback to the MailMerge object.
            doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

            // Prepare a simple data source.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("FirstName");
            table.Columns.Add("LastName");
            table.Columns.Add("Salary");
            table.Rows.Add("John", "Doe", 75000);
            table.Rows.Add("Jane", "Smith", 82000);

            // Execute the mail merge – the callback will format each field's text.
            doc.MailMerge.Execute(table);

            // Save the resulting document.
            doc.Save("CustomFormattedMailMerge.docx");
        }
    }
}
