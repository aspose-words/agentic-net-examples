using System;
using Aspose.Words;

namespace AsposeWordsExtensionDemo
{
    // Extension methods must be defined in a static, non‑generic class.
    public static class NumericExtensions
    {
        /// <summary>
        /// Returns true if the amount exceeds the specified limit.
        /// </summary>
        /// <param name="amount">The monetary amount to evaluate.</param>
        /// <param name="limit">The threshold limit.</param>
        /// <returns>True when amount &gt; limit; otherwise false.</returns>
        public static bool IsHighValue(this double amount, double limit)
        {
            return amount > limit;
        }
    }

    class Program
    {
        static void Main()
        {
            // Sample data.
            double amount = 1523.75;
            double limit = 1000.00;

            // Use the custom extension method.
            bool isHigh = amount.IsHighValue(limit);

            // Create a new Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a sentence that includes the amount and the evaluation result.
            builder.Writeln($"The amount is {amount:C2}.");
            builder.Writeln(isHigh
                ? $"It exceeds the limit of {limit:C2}."
                : $"It does not exceed the limit of {limit:C2}.");

            // Save the document.
            doc.Save("IsHighValueDemo.docx");
        }
    }
}
