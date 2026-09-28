using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple", Quantity = 1.5 },
                new Product { Name = "Banana", Quantity = 2.25 },
                new Product { Name = "Cherry", Quantity = 0.75 }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Product Report");
        builder.Writeln("----------------");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Quantity (fraction): <<[p.QuantityFraction]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        template.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public double Quantity { get; set; }

    // Returns the quantity formatted as a simple fraction (e.g., 1 1/2).
    public string QuantityFraction => ConvertToFraction(Quantity);

    private static string ConvertToFraction(double value)
    {
        // Simple conversion using a denominator of 4 (quarters). Adjust as needed.
        const int denominator = 4;
        int whole = (int)Math.Floor(value);
        double fractionalPart = value - whole;
        int numerator = (int)Math.Round(fractionalPart * denominator);

        // Reduce fraction if possible.
        int gcd = Gcd(numerator, denominator);
        numerator /= gcd;
        int reducedDenominator = denominator / gcd;

        if (numerator == 0)
            return whole.ToString();

        if (whole == 0)
            return $"{numerator}/{reducedDenominator}";

        return $"{whole} {numerator}/{reducedDenominator}";
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
