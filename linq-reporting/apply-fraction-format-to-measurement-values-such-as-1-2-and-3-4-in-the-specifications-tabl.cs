using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Specification
{
    public string Name { get; set; } = string.Empty;
    public double Measurement { get; set; }

    // Returns the measurement formatted as a simple fraction (e.g., 0.5 -> "1/2").
    public string MeasurementFraction
    {
        get
        {
            // Handle common fractions with a small tolerance.
            const double tolerance = 1e-6;

            // Define a list of denominators to try.
            int[] denominators = { 2, 4, 8, 16, 32, 64 };
            foreach (int d in denominators)
            {
                double numerator = Math.Round(Measurement * d);
                if (Math.Abs(Measurement - numerator / d) < tolerance && numerator > 0)
                {
                    // Reduce the fraction.
                    int n = (int)numerator;
                    int g = Gcd(n, d);
                    return $"{n / g}/{d / g}";
                }
            }

            // Fallback to decimal representation if no simple fraction found.
            return Measurement.ToString("0.##");
        }
    }

    // Greatest common divisor (Euclidean algorithm).
    private static int Gcd(int a, int b)
    {
        while (b != 0)
        {
            int temp = b;
            b = a % b;
            a = temp;
        }
        return a;
    }
}

public class ReportModel
{
    public List<Specification> Specs { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title.
        builder.Writeln("Specifications:");

        // Begin the foreach block.
        builder.Writeln("<<foreach [spec in Specs]>>");

        // Build a table with two columns: Name and Measurement (formatted as fraction).
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[spec.Name]>>");
        builder.InsertCell();
        // Use the custom property that returns the fraction string.
        builder.Writeln("<<[spec.MeasurementFraction]>>");
        builder.EndRow();

        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Specs = new List<Specification>
            {
                new Specification { Name = "Length", Measurement = 0.5 },   // 1/2
                new Specification { Name = "Width",  Measurement = 0.75 }   // 3/4
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
