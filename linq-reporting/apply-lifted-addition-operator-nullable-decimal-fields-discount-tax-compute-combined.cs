using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Nullable decimal fields that might come from some data source.
        decimal? discount = 12.5m;   // example value, can be null
        decimal? tax = null;        // example value, can be null

        // Lifted addition operator: if either operand is null the result is null.
        decimal? combined = discount + tax;

        // Write the values to the document for demonstration.
        builder.Writeln($"Discount: {(discount.HasValue ? discount.Value.ToString() : "null")}");
        builder.Writeln($"Tax: {(tax.HasValue ? tax.Value.ToString() : "null")}");
        builder.Writeln($"Combined (discount + tax): {(combined.HasValue ? combined.Value.ToString() : "null")}");

        // Save the document.
        doc.Save("CombinedValue.docx");
    }
}
