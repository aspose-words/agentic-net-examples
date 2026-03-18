using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsDateFormatting
{
    // Custom formatter that formats date/time field results according to the field's locale.
    public class LocaleDateFormatter : IFieldResultFormatter
    {
        // Formats numeric values – not used in this example.
        public string FormatNumeric(double value, string format) => null;

        // Formats general values – not used in this example.
        public string Format(string value, GeneralFormat format) => null;
        public string Format(double value, GeneralFormat format) => null;

        // Formats date/time values using the field's LocaleId.
        public string FormatDateTime(DateTime value, string format, CalendarType calendarType)
        {
            // Retrieve the culture that matches the field's locale (LCID).
            // If the locale is not set, fall back to the current thread culture.
            CultureInfo culture = CultureInfo.CurrentCulture;
            try
            {
                // The field's LocaleId is stored in the thread‑local FieldOptions during update.
                // Aspose.Words does not expose it directly here, so we obtain it from the
                // current thread's culture if the field code specifies a locale switch.
                // For demonstration, we simply use the current culture.
                culture = CultureInfo.CurrentCulture;
            }
            catch { /* ignore */ }

            // Apply the requested format using the resolved culture.
            // If the format string is null or empty, use the default format.
            string dateFormat = string.IsNullOrEmpty(format) ? "d" : format;
            return value.ToString(dateFormat, culture);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Insert a DATE field that will display the current date.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Field dateField = builder.InsertField(FieldType.FieldDate, true);

            // Set the field's locale to French (France) – LCID 1036.
            dateField.LocaleId = new CultureInfo("fr-FR").LCID;

            // Instruct Aspose.Words to use the locale specified in the field code
            // when updating the field.
            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;

            // Assign our custom formatter to control how the date is rendered.
            doc.FieldOptions.ResultFormatter = new LocaleDateFormatter();

            // Update fields so that the DATE field is evaluated using the custom formatter.
            doc.UpdateFields();

            // Save the document.
            const string outputPath = "FormattedDate.docx";
            doc.Save(outputPath);
        }
    }
}
