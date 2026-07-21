using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeMissingFieldDemo
{
    // Implements custom handling for missing merge fields.
    public class MissingFieldHandler : IFieldMergingCallback
    {
        // Called for each merge field encountered during mail merge.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // If the data source does not contain a value for this field,
            // Aspose.Words supplies a null value. Replace it with a clear placeholder.
            if (args.FieldValue == null || args.FieldValue == DBNull.Value)
            {
                // You could also log the missing field, throw an exception, etc.
                args.Text = $"[Missing: {args.DocumentFieldName}]";
            }
        }

        // No image handling required for this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a simple document with three merge fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("First Name: ");
            builder.InsertField("MERGEFIELD FirstName");
            builder.Writeln();

            builder.Write("Last Name: ");
            builder.InsertField("MERGEFIELD LastName");
            builder.Writeln();

            builder.Write("Address: ");
            builder.InsertField("MERGEFIELD Address");
            builder.Writeln();

            // Prepare a data source that intentionally lacks "LastName" and "Address".
            DataTable table = new DataTable("Customers");
            table.Columns.Add("FirstName");
            table.Rows.Add("John");

            // Attach the custom missing‑field handler.
            doc.MailMerge.FieldMergingCallback = new MissingFieldHandler();

            // Perform the mail merge.
            doc.MailMerge.Execute(table);

            // Save the result to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedResult.docx");
            doc.Save(outputPath);
        }
    }
}
