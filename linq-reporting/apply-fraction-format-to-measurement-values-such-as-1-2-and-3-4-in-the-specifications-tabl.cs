using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class SpecificationItem
{
    public string Name { get; set; } = string.Empty;
    public double Measurement { get; set; }

    // Returns a simple fractional representation of the measurement.
    // For the purpose of this example we limit the denominator to 8.
    public string FractionMeasurement
    {
        get
        {
            // Convert the double to a fraction with denominator up to 8.
            // This is a naive implementation sufficient for the sample data.
            const int maxDenominator = 8;
            double value = Measurement;
            int bestNumerator = 0;
            int bestDenominator = 1;
            double bestError = double.MaxValue;

            for (int denom = 1; denom <= maxDenominator; denom++)
            {
                int numer = (int)Math.Round(value * denom);
                double error = Math.Abs(value - (double)numer / denom);
                if (error < bestError)
                {
                    bestError = error;
                    bestNumerator = numer;
                    bestDenominator = denom;
                }
            }

            // If the fraction is an integer, just return the integer.
            if (bestDenominator == 1)
                return bestNumerator.ToString();

            return $"{bestNumerator}/{bestDenominator}";
        }
    }
}

public class ReportModel
{
    public List<SpecificationItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            Items = new List<SpecificationItem>
            {
                new() { Name = "Length", Measurement = 0.5 },   // 1/2
                new() { Name = "Width",  Measurement = 0.75 }, // 3/4
                new() { Name = "Height", Measurement = 0.3333 } // approx 1/3
            }
        };

        // Create the template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Specifications");
        builder.Writeln("<<foreach [item in Items]>>");

        // Table header.
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Measurement");
        builder.EndRow();

        // Table row bound to the data source.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        // Use the custom property that already contains the fraction string.
        builder.Writeln("<<[item.FractionMeasurement]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("SpecificationsReport.docx");
    }
}
