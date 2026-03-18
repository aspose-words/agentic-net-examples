using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Assign a custom result formatter that formats numeric field results as currency
        // using a .NET composite format string.
        doc.FieldOptions.ResultFormatter = new CurrencyResultFormatter();

        // Insert a numeric field. The field itself does not contain a format switch;
        // the formatter will apply the custom currency format.
        builder.InsertField("= 1234.56");

        // Update all fields so the formatter is invoked.
        doc.UpdateFields();

        // Save the document.
        doc.Save("CurrencyFormatted.docx");
    }

    // Custom formatter implementing IFieldResultFormatter.
    // Only FormatNumeric is overridden; other methods return null to keep default behavior.
    private class CurrencyResultFormatter : IFieldResultFormatter
    {
        public string FormatDateTime(DateTime value, string format, CalendarType calendarType) => null;

        public string Format(string value, GeneralFormat format) => null;

        public string Format(double value, GeneralFormat format) => null;

        // Apply a .NET format string to render the value as currency.
        // The format string "\"$\"{0:#,##0.00}" produces output like "$1,234.56".
        public string FormatNumeric(double value, string format)
        {
            return string.Format(CultureInfo.InvariantCulture, "\"$\"{0:#,##0.00}", value);
        }
    }
}
