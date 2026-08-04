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

        // Insert three merge fields: FirstName, LastName and Email.
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Write(" - ");
        builder.InsertField("MERGEFIELD Email");

        // Register a callback that supplies a default value for missing fields.
        doc.MailMerge.FieldMergingCallback = new MissingFieldHandler();

        // Build a data table that deliberately omits the "Email" column.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Rows.Add("John", "Doe");
        table.Rows.Add("Jane", "Smith");

        // Execute the mail merge. The callback will be invoked for the Email field.
        doc.MailMerge.Execute(table);

        // Save the merged document.
        doc.Save("MergedOutput.docx");
    }

    // Callback that handles missing fields during mail merge.
    private class MissingFieldHandler : IFieldMergingCallback
    {
        // Called for each merge field encountered.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // If the data source does not contain a value for this field, provide a placeholder.
            if (args.FieldValue == null || args.FieldValue == DBNull.Value)
            {
                // Setting Text replaces the field content with the specified string.
                args.Text = "[Missing]";
            }
        }

        // No image handling needed for this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Intentionally left blank.
        }
    }
}
