using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Specification
{
    public string Name { get; set; } = string.Empty;
    public double Measurement { get; set; }

    // Returns a simple fraction representation of the measurement.
    // This avoids using an unsupported format switch.
    public string FractionString
    {
        get
        {
            // Define a maximum denominator to search for a suitable fraction.
            const int maxDenominator = 16;
            const double tolerance = 1e-6;

            // Handle zero explicitly.
            if (Math.Abs(Measurement) < tolerance)
                return "0";

            // Work with absolute value for sign handling.
            double absValue = Math.Abs(Measurement);
            int sign = Measurement < 0 ? -1 : 1;

            // Find the best denominator.
            for (int denominator = 1; denominator <= maxDenominator; denominator++)
            {
                double numeratorExact = absValue * denominator;
                int numerator = (int)Math.Round(numeratorExact);
                if (Math.Abs(numeratorExact - numerator) < tolerance)
                {
                    // Reduce the fraction.
                    int gcd = Gcd(numerator, denominator);
                    numerator /= gcd;
                    denominator /= gcd;

                    // Apply sign to the numerator.
                    numerator *= sign;

                    return $"{numerator}/{denominator}";
                }
            }

            // Fallback: return the decimal value if no simple fraction found.
            return Measurement.ToString("0.###");
        }
    }

    // Euclidean algorithm for greatest common divisor.
    private static int Gcd(int a, int b)
    {
        a = Math.Abs(a);
        b = Math.Abs(b);
        while (b != 0)
        {
            int temp = b;
            b = a % b;
            a = temp;
        }
        return a == 0 ? 1 : a;
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
                new Specification { Name = "Length", Measurement = 0.5 },   // 1/2
                new Specification { Name = "Width",  Measurement = 0.75 }, // 3/4
                new Specification { Name = "Height", Measurement = 0.125 } // 1/8
            }
        };

        // Create a blank document and a builder.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Specifications Report");
        builder.Writeln();

        // Begin the foreach block.
        builder.Writeln("<<foreach [spec in Specs]>>");

        // Build a table for each specification.
        Table table = builder.StartTable();

        // Header row (once per foreach iteration).
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[spec.Name]>>");
        builder.InsertCell();
        // Use the custom fraction string property.
        builder.Writeln("<<[spec.FractionString]>>");
        builder.EndRow();

        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
