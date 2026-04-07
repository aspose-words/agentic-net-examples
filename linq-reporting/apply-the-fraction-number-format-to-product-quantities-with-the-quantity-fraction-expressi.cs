using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = string.Empty;
    public double Quantity { get; set; }

    // Returns the quantity formatted as a fraction (e.g., 1.5 -> "1 1/2").
    public string FormattedQuantity => ToFraction(Quantity);

    private static string ToFraction(double value)
    {
        // Simple conversion: separate whole part and fractional part (up to 1/8 precision).
        int whole = (int)Math.Floor(value);
        double fraction = value - whole;

        // Define common fractions.
        var fractions = new (double value, string text)[]
        {
            (0.0, ""),
            (0.125, "1/8"),
            (0.25, "1/4"),
            (0.375, "3/8"),
            (0.5, "1/2"),
            (0.625, "5/8"),
            (0.75, "3/4"),
            (0.875, "7/8")
        };

        // Find the closest fraction.
        string fractionText = "";
        double minDiff = double.MaxValue;
        foreach (var (fracValue, text) in fractions)
        {
            double diff = Math.Abs(fracValue - fraction);
            if (diff < minDiff)
            {
                minDiff = diff;
                fractionText = text;
            }
        }

        if (whole == 0 && !string.IsNullOrEmpty(fractionText))
            return fractionText;
        if (whole != 0 && string.IsNullOrEmpty(fractionText))
            return whole.ToString();
        return $"{whole} {fractionText}".Trim();
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
        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apples", Quantity = 1.5 },
                new Product { Name = "Oranges", Quantity = 2.75 }
            }
        };

        // Create a template document programmatically.
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [product in Products]>>");
        builder.Writeln("Name: <<[product.Name]>>, Quantity: <<[product.FormattedQuantity]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template and build the report.
        var template = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}
