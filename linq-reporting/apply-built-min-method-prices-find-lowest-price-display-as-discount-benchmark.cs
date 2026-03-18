using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;

class DiscountBenchmark
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample price data.
        double[] prices = { 199.99, 149.50, 179.75, 129.99, 159.00 };

        // Use the built‑in LINQ Min method to find the lowest price.
        double minPrice = prices.Min();

        // Write the discount benchmark into the document.
        builder.Writeln($"Discount benchmark (lowest price): {minPrice:C}");

        // Save the document.
        doc.Save("DiscountBenchmark.docx");
    }
}
