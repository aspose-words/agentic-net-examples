using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public double Quantity { get; set; }

    // Returns the quantity formatted as a fraction (e.g., 1.5 -> "1 1/2").
    public string FormattedQuantity => ToFraction(Quantity);

    private static string ToFraction(double value)
    {
        // Separate whole part and fractional part.
        int whole = (int)Math.Floor(value);
        double fraction = value - whole;

        // Define a set of common denominators.
        int[] denominators = { 2, 3, 4, 8, 16, 32, 64 };
        int bestDenominator = 1;
        int bestNumerator = 0;
        double minDiff = double.MaxValue;

        foreach (int d in denominators)
        {
            int n = (int)Math.Round(fraction * d);
            double diff = Math.Abs(fraction - (double)n / d);
            if (diff < minDiff)
            {
                minDiff = diff;
                bestDenominator = d;
                bestNumerator = n;
            }
        }

        // If the fractional part is effectively zero, return only the whole number.
        if (bestNumerator == 0)
            return whole.ToString(CultureInfo.InvariantCulture);

        // If there is no whole part, return only the fraction.
        if (whole == 0)
            return $"{bestNumerator}/{bestDenominator}";

        return $"{whole} {bestNumerator}/{bestDenominator}";
    }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings Aspose might need.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apples",  Quantity = 1.5 },
                new Product { Name = "Oranges", Quantity = 2.25 },
                new Product { Name = "Bananas", Quantity = 0.75 }
            }
        };

        // Create a template document.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Product: <<[p.Name]>>");
        // Use the custom formatted property instead of the unsupported :fraction switch.
        builder.Writeln("Qty: <<[p.FormattedQuantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, shown for completeness).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}
