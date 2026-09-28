using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with two merge fields.
        builder.Writeln("Hello <<Name>>, today is <<Date>>.");

        // Set up a custom field merging callback to format inserted text.
        doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

        // Define field names and corresponding values.
        string[] fieldNames = { "Name", "Date" };
        object[] fieldValues = { "John Doe", DateTime.Now };

        // Execute mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the result to a file.
        doc.Save("MergedOutput.docx");

        // Indicate completion.
        Console.WriteLine("Mail merge completed. Output saved to MergedOutput.docx");
    }

    // Custom callback to control how each merge field value is inserted.
    private class CustomFieldMergingCallback : IFieldMergingCallback
    {
        // Called for each merge field during mail merge.
        public void FieldMerging(FieldMergingArgs e)
        {
            // Format the value based on the field name.
            if (e.FieldName.Equals("Name", StringComparison.OrdinalIgnoreCase))
            {
                // Example: convert name to uppercase.
                e.Text = ((string)e.FieldValue).ToUpperInvariant();
            }
            else if (e.FieldName.Equals("Date", StringComparison.OrdinalIgnoreCase))
            {
                // Example: format date as "Month day, year".
                DateTime date = (DateTime)e.FieldValue;
                e.Text = date.ToString("MMMM dd, yyyy");
            }
            else
            {
                // Default handling for any other fields.
                e.Text = e.FieldValue?.ToString() ?? string.Empty;
            }
        }

        // Not used in this example but required by the interface.
        public void ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
