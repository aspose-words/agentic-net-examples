using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some MERGEFIELDs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Write(" is ");
        builder.InsertField("MERGEFIELD Age");
        builder.Writeln(" years old.");

        // Use a custom callback to handle fields that are missing in the data source.
        doc.MailMerge.FieldMergingCallback = new MissingFieldHandler();

        // Prepare a data source that intentionally lacks the "Age" column.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Rows.Add("John", "Doe");
        table.Rows.Add("Jane", "Smith");

        // Execute the mail merge. The callback will supply a default value for "Age".
        doc.MailMerge.Execute(table);

        // Save the merged document.
        doc.Save("MergedDocument.docx");
    }

    // Callback that supplies a default value for missing fields or removes them.
    private class MissingFieldHandler : IFieldMergingCallback
    {
        // Called for each merge field encountered during mail merge.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // If the data source does not contain a value for this field, args.FieldValue will be null.
            if (args.FieldValue == null)
            {
                // Provide a default value for the "Age" field.
                if (args.DocumentFieldName.Equals("Age", StringComparison.OrdinalIgnoreCase))
                {
                    args.Text = "N/A";
                }
                else
                {
                    // For any other missing field, remove the field from the document.
                    args.Text = string.Empty;
                }
            }
        }

        // No image handling required for this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
