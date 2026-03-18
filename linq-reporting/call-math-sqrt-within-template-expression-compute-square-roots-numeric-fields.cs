using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a numeric formula field. The field result will be 25.
        builder.InsertField(" = 25 \\# 0", null);

        // Attach a custom result formatter that replaces the numeric result with its square root.
        doc.FieldOptions.ResultFormatter = new SqrtResultFormatter();

        // Update fields so the formatter is applied.
        doc.UpdateFields();

        // Save the document.
        doc.Save("SqrtField.docx");
    }

    // Custom formatter that computes the square root of numeric field results.
    private class SqrtResultFormatter : IFieldResultFormatter
    {
        // Called for numeric fields. Returns the square root of the original value.
        public string FormatNumeric(double value, string format)
        {
            double sqrt = Math.Sqrt(value);
            // Preserve the original numeric format if needed; here we simply return the raw value.
            return sqrt.ToString();
        }

        // Default handling for date/time fields (no custom formatting).
        public string FormatDateTime(DateTime value, string format, CalendarType calendarType) => null;

        // Default handling for general string formatting (no custom formatting).
        public string Format(string value, GeneralFormat format) => null;

        // Default handling for general numeric formatting (no custom formatting).
        public string Format(double value, GeneralFormat format) => null;
    }
}
