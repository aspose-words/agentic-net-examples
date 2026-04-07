using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Specification
{
    public string Name { get; set; } = "";
    public double Value { get; set; }

    // Returns a simple fraction representation for common values (e.g., 0.5 -> 1/2, 0.75 -> 3/4)
    public string Fraction
    {
        get
        {
            // Round to nearest 1/16 to avoid floating‑point noise.
            const int denominator = 16;
            int numerator = (int)Math.Round(Value * denominator);
            // Reduce the fraction.
            int gcd = Gcd(numerator, denominator);
            numerator /= gcd;
            int reducedDenominator = denominator / gcd;

            // If denominator becomes 1, return whole number.
            return reducedDenominator == 1 ? numerator.ToString() : $"{numerator}/{reducedDenominator}";
        }
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
    public List<Specification> Specs { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and final report.
        const string templatePath = "SpecificationTemplate.docx";
        const string reportPath = "SpecificationReport.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Product Specifications");
        builder.Writeln();

        // Begin foreach loop over Specs collection.
        builder.Writeln("<<foreach [spec in Specs]>>");

        // Build a simple two‑column table: Name | Measurement.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Measurement");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[spec.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[spec.Fraction]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Prepare sample data.
        // -------------------------------------------------
        ReportModel model = new ReportModel();
        model.Specs.Add(new Specification { Name = "Length", Value = 0.5 });   // 1/2
        model.Specs.Add(new Specification { Name = "Width", Value = 0.75 });   // 3/4
        model.Specs.Add(new Specification { Name = "Height", Value = 1.0 });   // 1

        // -------------------------------------------------
        // 3. Build the report using ReportingEngine.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
