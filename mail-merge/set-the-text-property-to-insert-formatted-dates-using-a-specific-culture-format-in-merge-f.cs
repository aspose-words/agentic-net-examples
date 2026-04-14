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

        // Insert a MERGEFIELD with a date format switch.
        // The switch \\@ defines the date format that would be used if the field were updated automatically.
        builder.InsertField("MERGEFIELD FormattedDate \\@ \"dddd, d MMMM yyyy\"");

        // Assign a custom field merging callback to control how the merge value is inserted.
        doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

        // Perform the mail merge with a single date value.
        DateTime mergeDate = new DateTime(2023, 12, 25);
        doc.MailMerge.Execute(new[] { "FormattedDate" }, new object[] { mergeDate });

        // Save the result to a file in the current directory.
        doc.Save("FormattedDateMerge.docx");
    }

    // Custom callback that formats DateTime values using a specific culture before insertion.
    private class CustomFieldMergingCallback : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            // Ensure the field value is a DateTime.
            if (args.FieldValue is DateTime dateValue)
            {
                // Define the desired culture (e.g., French - France).
                CultureInfo targetCulture = new CultureInfo("fr-FR");

                // Use the same format as defined in the field switch.
                // Here we use a custom format string; you can adjust as needed.
                string formattedDate = dateValue.ToString("dddd, d MMMM yyyy", targetCulture);

                // Set the Text property to insert the formatted date.
                args.Text = formattedDate;
            }
            else
            {
                // For non‑date fields, fall back to the default behavior.
                args.Text = args.FieldValue?.ToString();
            }
        }

        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
