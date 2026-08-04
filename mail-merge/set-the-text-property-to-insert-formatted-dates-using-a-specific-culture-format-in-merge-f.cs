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

        // Insert a MERGEFIELD that expects a DateTime value.
        // The field includes a date format switch (\\@) which will be ignored because we set the Text property manually.
        builder.InsertField("MERGEFIELD MyDate \\@ \"dddd, d MMMM yyyy\"");

        // Assign a custom field merging callback to format dates.
        doc.MailMerge.FieldMergingCallback = new DateFormattingCallback();

        // Execute the mail merge with a single date value.
        DateTime mergeDate = new DateTime(2023, 12, 25);
        doc.MailMerge.Execute(new[] { "MyDate" }, new object[] { mergeDate });

        // Save the result to disk.
        doc.Save("FormattedDateMerge.docx");
    }

    // Custom callback that formats DateTime values using a specific culture and assigns the result to the Text property.
    private class DateFormattingCallback : IFieldMergingCallback
    {
        // This method is called for each merge field during the mail merge operation.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Check if the field value is a DateTime.
            if (args.FieldValue is DateTime dateValue)
            {
                // Define the culture you want to use for formatting (e.g., German - Germany).
                CultureInfo culture = new CultureInfo("de-DE");

                // Define the desired date format.
                string format = "dddd, d MMMM yyyy";

                // Format the date using the specified culture.
                string formattedDate = dateValue.ToString(format, culture);

                // Set the Text property so that the formatted string is inserted into the document.
                args.Text = formattedDate;
            }
            else
            {
                // For non‑date fields, let the default behavior occur.
                args.Text = null;
            }
        }

        // This method is required by the interface but is not needed for this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
