using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Spec
{
    public int Index { get; set; }
    public double Value { get; set; }

    // Returns the value formatted as a fraction (e.g., 1/2, 3/4).
    public string Fraction => ConvertToFraction(Value);

    private static string ConvertToFraction(double value)
    {
        // Simple handling for common fractions.
        if (Math.Abs(value - 0.5) < 0.0001) return "1/2";
        if (Math.Abs(value - 0.75) < 0.0001) return "3/4";

        // Generic conversion with a denominator up to 100.
        const int maxDenominator = 100;
        int denominator = maxDenominator;
        int numerator = (int)Math.Round(value * denominator);
        int gcd = Gcd(numerator, denominator);
        numerator /= gcd;
        denominator /= gcd;
        return $"{numerator}/{denominator}";
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
    public List<Spec> Specs { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel();
        model.Specs.Add(new Spec { Index = 1, Value = 0.5 });
        model.Specs.Add(new Spec { Index = 2, Value = 0.75 });
        model.Specs.Add(new Spec { Index = 3, Value = 0.3333 });

        // ---------- Create the template document ----------
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Specifications Table:");
        // Begin the foreach block before the table.
        builder.Writeln("<<foreach [spec in Specs]>>");

        // Build the table that will be repeated for each Spec.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Value (Fraction)");
        builder.EndRow();

        // Data row – will be repeated.
        builder.InsertCell();
        builder.Writeln("<<[spec.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[spec.Fraction]>>");
        builder.EndRow();

        // Finish the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // ---------- Load the template and build the report ----------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        doc.Save("Report.docx");
    }
}
