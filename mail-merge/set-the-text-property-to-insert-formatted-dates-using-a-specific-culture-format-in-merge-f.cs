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

        // Insert a MERGEFIELD that expects a date value and includes a date format switch.
        // The format switch (\\@) will be ignored because we will supply the formatted text ourselves.
        builder.InsertField("MERGEFIELD OrderDate \\@ \"dddd, d MMMM yyyy\"");

        // Register a field merging callback that formats DateTime values using a specific culture.
        doc.MailMerge.FieldMergingCallback = new DateFormattingCallback();

        // Perform the mail merge with a single date value.
        DateTime orderDate = new DateTime(2023, 12, 15);
        doc.MailMerge.Execute(new[] { "OrderDate" }, new object[] { orderDate });

        // Save the result to a file.
        doc.Save("FormattedDates.docx");
    }

    // Implements IFieldMergingCallback to control the text inserted for merge fields.
    private class DateFormattingCallback : IFieldMergingCallback
    {
        // Called for each MERGEFIELD encountered during mail merge.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Check that the field value is a DateTime.
            if (args.FieldValue is DateTime dateValue)
            {
                // Define the culture you want to use for formatting (e.g., German).
                CultureInfo culture = new CultureInfo("de-DE");

                // Apply the same format string that was used in the field's \\@ switch.
                // This ensures the output matches the expected pattern.
                string format = "dddd, d MMMM yyyy";

                // Format the date using the specified culture.
                string formattedDate = dateValue.ToString(format, culture);

                // Set the Text property to the formatted string.
                args.Text = formattedDate;
            }
        }

        // Required by the interface but not needed for this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
