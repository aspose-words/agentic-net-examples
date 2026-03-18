using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Create a document containing several MERGEFIELDs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertField(" MERGEFIELD FirstName ");
        builder.Write(" ");
        builder.InsertField(" MERGEFIELD LastName ");
        builder.Write(" ");
        builder.InsertField(" MERGEFIELD Age "); // This field will be missing in the data source.

        // Build a data source that does NOT contain the "Age" column.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Rows.Add("John", "Doe");
        table.Rows.Add("Jane", "Smith");

        // Attach a callback that will handle missing fields during the merge.
        doc.MailMerge.FieldMergingCallback = new MissingFieldHandler();

        // Perform the mail merge.
        doc.MailMerge.Execute(table);

        // Save the merged document.
        doc.Save("MissingFieldHandled.docx");
    }

    // Implements IFieldMergingCallback to intercept each field during mail merge.
    private class MissingFieldHandler : IFieldMergingCallback
    {
        // Called for every simple MERGEFIELD.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // If the data source does not provide a value, FieldValue will be null.
            // Replace the field with a custom warning text.
            if (args.FieldValue == null)
            {
                args.Text = $"[Missing field: {args.DocumentFieldName}]";
            }
            // If a value exists, leave args.Text as null so the engine inserts the value.
        }

        // No special handling for image fields in this scenario.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Intentionally left blank.
        }
    }
}
