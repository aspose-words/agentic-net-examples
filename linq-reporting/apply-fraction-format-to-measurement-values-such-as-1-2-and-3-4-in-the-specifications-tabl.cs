using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Specification
{
    public string Name { get; set; } = "";
    public double Value { get; set; }

    // Returns a simple fraction representation for demonstration purposes.
    public string Fraction => ConvertToFraction(Value);

    private static string ConvertToFraction(double value)
    {
        // Very basic conversion for common fractions used in the example.
        if (Math.Abs(value - 0.5) < 0.0001) return "1/2";
        if (Math.Abs(value - 0.75) < 0.0001) return "3/4";
        if (Math.Abs(value - 0.25) < 0.0001) return "1/4";

        // Fallback to decimal string if not a known fraction.
        return value.ToString("0.##");
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
        // Sample data.
        var model = new ReportModel
        {
            Specs = new List<Specification>
            {
                new Specification { Name = "Length", Value = 0.5 },
                new Specification { Name = "Width",  Value = 0.75 },
                new Specification { Name = "Height", Value = 0.25 }
            }
        };

        // Create the template document.
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Specifications Report");
        builder.Writeln();

        // Begin foreach loop.
        builder.Writeln("<<foreach [spec in Specs]>>");

        // Table header (once, before the loop).
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Parameter");
        builder.InsertCell();
        builder.Writeln("Measurement");
        builder.EndRow();

        // Data rows – repeated for each spec.
        builder.InsertCell();
        builder.Writeln("<<[spec.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[spec.Fraction]>>");
        builder.EndRow();

        // End table and foreach.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
