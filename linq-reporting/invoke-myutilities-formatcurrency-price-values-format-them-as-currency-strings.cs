using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a numeric field. The field result will be a number (e.g., 1234.56).
        // The numeric format switch (\# 0.00) ensures the field evaluates to a decimal string.
        builder.InsertField("= 1234.56 \\# 0.00");

        // Assign a custom result formatter that uses MyUtilities.FormatCurrency
        // to format numeric field results as currency strings.
        doc.FieldOptions.ResultFormatter = new CurrencyResultFormatter();

        // Update all fields so the custom formatter is applied.
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save("FormattedCurrency.docx");
    }

    // Custom formatter implementing IFieldResultFormatter.
    // Only FormatNumeric is overridden to apply currency formatting.
    private class CurrencyResultFormatter : IFieldResultFormatter
    {
        // Called for numeric fields (the \# switch).
        public string FormatNumeric(double value, string format)
        {
            // Delegate formatting to the external utility.
            return MyUtilities.FormatCurrency(value);
        }

        // No custom handling for date/time fields.
        public string FormatDateTime(DateTime value, string format, CalendarType calendarType)
        {
            return null; // Use default formatting.
        }

        // No custom handling for general text formatting.
        public string Format(string value, GeneralFormat format)
        {
            return null; // Use default formatting.
        }

        // No custom handling for general numeric formatting.
        public string Format(double value, GeneralFormat format)
        {
            return null; // Use default formatting.
        }
    }
}

// External utility class that formats a double as a currency string.
public static class MyUtilities
{
    public static string FormatCurrency(double value)
    {
        // Example: format using invariant culture with a dollar sign and two decimal places.
        return string.Format(CultureInfo.InvariantCulture, "\"$\"{0:N2}", value);
    }
}
