using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = string.Empty;
    public double Quantity { get; set; }

    // Returns the quantity formatted as a simple fraction (e.g., "1 1/2").
    public string FractionQuantity => ConvertToFraction(Quantity);

    private static string ConvertToFraction(double value)
    {
        // Separate whole number part.
        int whole = (int)Math.Floor(value);
        double fractional = value - whole;

        // Use a denominator of 16 for reasonable precision.
        const int maxDenominator = 16;
        int numerator = (int)Math.Round(fractional * maxDenominator);
        int denominator = maxDenominator;

        // Reduce the fraction.
        int gcd = Gcd(numerator, denominator);
        numerator /= gcd;
        denominator /= gcd;

        // Build the string representation.
        if (numerator == 0)
            return whole.ToString();
        if (whole == 0)
            return $"{numerator}/{denominator}";
        return $"{whole} {numerator}/{denominator}";
    }

    private static int Gcd(int a, int b)
    {
        while (b != 0)
        {
            int temp = b;
            b = a % b;
            a = temp;
        }
        return Math.Abs(a);
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
                new Product { Name = "Apples",  Quantity = 1.5 },
                new Product { Name = "Oranges", Quantity = 2.25 },
                new Product { Name = "Bananas", Quantity = 0.75 }
            }
        };

        // Create a blank document and build the template.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        // Use the custom FractionQuantity property instead of an unsupported format tag.
        builder.Writeln("Quantity: <<[p.FractionQuantity]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("ProductReport.docx");
    }
}
