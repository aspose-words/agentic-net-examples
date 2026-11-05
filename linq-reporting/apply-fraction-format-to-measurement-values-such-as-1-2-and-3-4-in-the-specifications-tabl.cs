using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class Specification
{
    public string Name { get; set; } = "";
    public double Value { get; set; }

    // Returns the value formatted as a simple fraction (e.g., 1/2, 3/4).
    public string Fraction => ToFraction(Value);

    private static string ToFraction(double value)
    {
        // Handle common fractions up to denominator 16.
        const int maxDenominator = 16;
        int sign = Math.Sign(value);
        value = Math.Abs(value);
        int bestNumerator = 0;
        int bestDenominator = 1;
        double bestError = double.MaxValue;

        for (int denominator = 1; denominator <= maxDenominator; denominator++)
        {
            int numerator = (int)Math.Round(value * denominator);
            double error = Math.Abs(value - (double)numerator / denominator);
            if (error < bestError)
            {
                bestError = error;
                bestNumerator = numerator;
                bestDenominator = denominator;
            }
        }

        if (bestDenominator == 1)
            return (sign * bestNumerator).ToString();

        return $"{(sign * bestNumerator)}/{bestDenominator}";
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
        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Specifications");
        builder.Writeln("<<foreach [spec in Specs]>>");

        // Start a table for each specification (as per original logic).
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Specification");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        // Data row – values will be filled by the engine.
        builder.InsertCell();
        builder.Writeln("<<[spec.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[spec.Fraction]>>");
        builder.EndRow();

        // End the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Sample data.
        var model = new ReportModel
        {
            Specs = new List<Specification>
            {
                new Specification { Name = "Length", Value = 0.5 },   // 1/2
                new Specification { Name = "Width",  Value = 0.75 }, // 3/4
                new Specification { Name = "Height", Value = 1.0 }   // 1
            }
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
