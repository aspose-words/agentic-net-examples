using System;
using System.Data;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD with a date format switch.
        // The field will display the date according to the format we provide in the callback.
        builder.InsertField("MERGEFIELD Date \\@ \"dddd, d MMMM yyyy\"");

        // Prepare a data source with a single DateTime value.
        DataTable table = new DataTable("Data");
        table.Columns.Add("Date", typeof(DateTime));
        table.Rows.Add(new DateTime(2020, 1, 1));

        // Assign a custom field merging callback that formats the date using a specific culture.
        doc.MailMerge.FieldMergingCallback = new DateFormattingCallback();

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Output the merged result to the console.
        Console.WriteLine(doc.Range.Text.Trim());
    }

    // Custom callback that formats DateTime values using the German culture.
    private class DateFormattingCallback : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            // Ensure the field value is a DateTime.
            if (args.FieldValue is DateTime dateValue)
            {
                // Use German culture for formatting.
                CultureInfo germanCulture = new CultureInfo("de-DE");
                // Apply the same format as defined in the field switch.
                string formatted = dateValue.ToString("dddd, d MMMM yyyy", germanCulture);
                // Set the Text property to insert the formatted date.
                args.Text = formatted;
            }
            else
            {
                // For non‑DateTime fields, fall back to the default behavior.
                args.Text = null;
            }
        }

        // Required by the interface but not used in this example.
        public void ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
