using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FinancialRoundedValuesExample
{
    // Custom formatter that uses System.Math.Round to round numeric results.
    class MathRoundResultFormatter : IFieldResultFormatter
    {
        public string FormatNumeric(double value, string format)
        {
            // Round to two decimal places using System.Math.
            double rounded = Math.Round(value, 2, MidpointRounding.AwayFromZero);
            // Apply the original numeric format if one exists.
            return !string.IsNullOrEmpty(format) ? rounded.ToString(format) : rounded.ToString();
        }

        public string FormatDateTime(DateTime value, string format, CalendarType calendarType) => null;
        public string Format(string value, GeneralFormat format) => null;
        public string Format(double value, GeneralFormat format) => null;
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to work with the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a formula field that contains the raw numeric value.
            // The field code is "= 1234.5678". This will be evaluated by Word.
            Field field = builder.InsertField("= 1234.5678");

            // Apply a numeric format that rounds the result to two decimal places.
            // The format "#.00" forces Word to display exactly two digits after the decimal point.
            field.Format.NumericFormat = "#.00";

            // Update the field so that the calculation and formatting are applied.
            field.Update();

            // Assign the custom formatter to the document.
            doc.FieldOptions.ResultFormatter = new MathRoundResultFormatter();

            // Insert a second formula field that will be processed by the custom formatter.
            Field field2 = builder.InsertField("= 9876.54321");

            // Update the second field; the formatter will round it to two decimals.
            field2.Update();

            // Save the document to disk.
            doc.Save("FinancialRoundedValues.docx");
        }
    }
}
