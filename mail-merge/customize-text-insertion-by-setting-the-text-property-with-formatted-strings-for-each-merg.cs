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

        // Insert merge fields into the document.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.Write("You are ");
        builder.InsertField("MERGEFIELD Age", "<Age>");
        builder.Writeln(" years old.");

        // Prepare a data source.
        DataTable table = new DataTable("People");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Columns.Add("Age");
        table.Rows.Add("John", "Doe", 30);
        table.Rows.Add("Jane", "Smith", 25);

        // Assign a custom field merging callback to format the inserted text.
        doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

        // Perform the mail merge.
        doc.MailMerge.Execute(table);

        // Save the result.
        doc.Save("Result.docx");
    }

    // Custom callback that sets the Text property for each merge field.
    private class CustomFieldMergingCallback : IFieldMergingCallback
    {
        // Called for each merge field.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Example of custom formatting based on the field name.
            switch (args.DocumentFieldName)
            {
                case "FirstName":
                    // Uppercase the first name.
                    args.Text = args.FieldValue?.ToString().ToUpper();
                    break;
                case "LastName":
                    // Title case the last name.
                    args.Text = args.FieldValue?.ToString();
                    break;
                case "Age":
                    // Append a descriptive suffix.
                    args.Text = $"{args.FieldValue} years old";
                    break;
                default:
                    // Default insertion.
                    args.Text = args.FieldValue?.ToString();
                    break;
            }
        }

        // Required by the interface but not used in this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling needed.
        }
    }
}
