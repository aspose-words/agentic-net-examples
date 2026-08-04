using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert MERGEFIELDs that will be populated during mail merge.
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln();

        // Prepare a simple data source.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Rows.Add("John", "Doe");
        table.Rows.Add("Jane", "Smith");

        // Assign a custom callback that formats the inserted text.
        doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

        // Perform the mail merge.
        doc.MailMerge.Execute(table);

        // Save the merged document.
        doc.Save("MergedDocument.docx");
    }

    // Custom callback that sets the Text property for each merge field.
    private class CustomFieldMergingCallback : IFieldMergingCallback
    {
        // Called for each simple merge field.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Example format: "FirstName = John"
            string value = args.FieldValue?.ToString() ?? "<null>";
            args.Text = $"{args.FieldName} = {value}";
        }

        // Not handling image fields in this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No custom image handling required.
        }
    }
}
