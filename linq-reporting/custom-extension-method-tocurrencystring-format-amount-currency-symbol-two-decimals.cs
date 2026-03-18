using System;
using System.Globalization;
using Aspose.Words;

namespace AsposeWordsCurrencyExample
{
    // Extension methods for numeric types.
    public static class NumericExtensions
    {
        // Formats a numeric value as a currency string with a leading '$' and two decimal places.
        public static string ToCurrencyString(this double amount)
        {
            // Use invariant culture to ensure '.' as decimal separator.
            return string.Format(CultureInfo.InvariantCulture, "${0:0.00}", amount);
        }

        // Overload for decimal values.
        public static string ToCurrencyString(this decimal amount)
        {
            return string.Format(CultureInfo.InvariantCulture, "${0:0.00}", amount);
        }
    }

    class Program
    {
        static void Main()
        {
            // Sample amount to format.
            double amount = 1234.5;

            // Apply the custom extension method.
            string formattedAmount = amount.ToCurrencyString(); // "$1234.50"

            // Create a new Aspose.Words document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the formatted amount into the document.
            builder.Writeln($"Amount: {formattedAmount}");

            // Save the document.
            doc.Save("FormattedAmount.docx");
        }
    }
}
