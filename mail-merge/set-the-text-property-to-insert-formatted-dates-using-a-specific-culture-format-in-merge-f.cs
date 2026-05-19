using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD named "Date".
        builder.InsertField("MERGEFIELD Date", "<Date>");

        // Assign a custom field merging callback to format the date.
        doc.MailMerge.FieldMergingCallback = new DateFormattingCallback();

        // Prepare the data for the merge.
        string[] fieldNames = { "Date" };
        object[] fieldValues = { new DateTime(2020, 1, 1) };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the result to a file.
        doc.Save("FormattedDate.docx");
    }

    // Custom callback that sets the Text property with a culture‑specific formatted date.
    private class DateFormattingCallback : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            if (args.DocumentFieldName.Equals("Date", StringComparison.OrdinalIgnoreCase) &&
                args.FieldValue is DateTime date)
            {
                // Use German culture for formatting.
                CultureInfo germanCulture = new CultureInfo("de-DE");
                string formattedDate = date.ToString("dddd, d MMMM yyyy", germanCulture);
                args.Text = formattedDate;
            }
        }

        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
