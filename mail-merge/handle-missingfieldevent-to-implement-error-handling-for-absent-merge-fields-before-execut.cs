using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    // Custom callback to detect missing merge fields.
    private class MissingFieldHandler : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            // If the data source does not contain a value for the field, throw an exception.
            if (args.FieldValue == null)
                throw new InvalidOperationException($"Missing merge field: {args.FieldName}");
        }

        // Not used in this example but required by the interface.
        public void ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    public static void Main()
    {
        // Create a new document and build the template with merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln(",");

        // Attach the custom callback to handle missing fields.
        doc.MailMerge.FieldMergingCallback = new MissingFieldHandler();

        // Provide data that lacks the LastName field.
        string[] fieldNames = { "FirstName" };
        object[] fieldValues = { "John" };

        try
        {
            // Execute the mail merge.
            doc.MailMerge.Execute(fieldNames, fieldValues);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Mail merge failed: {ex.Message}");
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
