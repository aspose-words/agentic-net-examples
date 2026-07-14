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

        // Insert MERGEFIELDs for first and last name.
        builder.InsertField(" MERGEFIELD FirstName ");
        builder.Write(" ");
        builder.InsertField(" MERGEFIELD LastName ");
        builder.Writeln();

        // Prepare a data table with sample names.
        DataTable table = new DataTable("Names");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Rows.Add("John", "Doe");
        table.Rows.Add("Jane", "Smith");

        // Set a custom callback that writes the field value in bold.
        doc.MailMerge.FieldMergingCallback = new BoldFieldMergingCallback();

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        doc.Save("MergedBoldNames.docx");
    }

    // Custom callback that inserts the merge field value with bold formatting.
    private class BoldFieldMergingCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Move the builder to the location of the current merge field.
            DocumentBuilder builder = new DocumentBuilder(args.Document);
            builder.MoveToMergeField(args.DocumentFieldName);

            // Apply bold formatting and write the field value.
            builder.Font.Bold = true;
            builder.Write(args.FieldValue?.ToString() ?? string.Empty);

            // Suppress the default insertion of the field value.
            args.Text = string.Empty;
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
