using System;
using System.Data;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build the document: insert a MERGEFIELD that will receive a DateTime value.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The field name is "DateField". No format switch is added here – we will format the value in code.
        builder.InsertField("MERGEFIELD DateField", null);

        // Attach a custom field merging callback that formats the date using a specific culture.
        doc.MailMerge.FieldMergingCallback = new DateFormattingCallback();

        // Prepare the data source for the mail merge.
        DataTable table = new DataTable("Data");
        table.Columns.Add("DateField", typeof(DateTime));
        // Example date to be merged.
        table.Rows.Add(new DateTime(2023, 12, 25));

        // Execute the mail merge. The callback will set the Text property for the field.
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        doc.Save("FormattedDateMerge.docx");
    }

    // Custom callback that formats the merged value as a date string using a specific culture.
    private class DateFormattingCallback : IFieldMergingCallback
    {
        // This method is called for each MERGEFIELD during the mail merge.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Only handle the "DateField" merge field.
            if (args.DocumentFieldName.Equals("DateField", StringComparison.OrdinalIgnoreCase) &&
                args.FieldValue is DateTime dateValue)
            {
                // Define the culture you want to use for formatting (e.g., French - France).
                CultureInfo culture = new CultureInfo("fr-FR");

                // Define the desired date format pattern.
                string formatPattern = "dddd, d MMMM yyyy";

                // Set the Text property – this string will be inserted into the document
                // instead of the raw field value.
                args.Text = dateValue.ToString(formatPattern, culture);
            }
        }

        // No special handling for image fields.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
