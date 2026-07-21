using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeCustomText
{
    // Custom callback that formats the text inserted for each merge field.
    public class CustomFieldMergingCallback : IFieldMergingCallback
    {
        // This method is called for every MERGEFIELD during the mail merge.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Example formatting: "Value (Record #X)" where X is zero‑based record index.
            string formatted = $"{args.FieldValue} (Record #{args.RecordIndex + 1})";
            args.Text = formatted;
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

            // Insert two merge fields.
            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(".");

            // Assign the custom field merging callback.
            doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

            // Prepare a data source.
            DataTable table = new DataTable("People");
            table.Columns.Add("FirstName");
            table.Columns.Add("LastName");
            table.Rows.Add("John", "Doe");
            table.Rows.Add("Jane", "Smith");

            // Execute the mail merge using the DataTable.
            doc.MailMerge.Execute(table);

            // Save the result to a file in the current directory.
            doc.Save("MergedResult.docx");
        }
    }
}
