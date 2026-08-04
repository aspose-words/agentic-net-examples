using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model representing a financial record.
    public class FinancialRecord
    {
        public string Description { get; set; } = string.Empty;
        public double Amount { get; set; }
    }

    // Wrapper class that will be passed as the root data source.
    public class ReportModel
    {
        public List<FinancialRecord> Items { get; set; } = new();
    }

    public static void Main()
    {
        // 1. Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<FinancialRecord>
            {
                new FinancialRecord { Description = "Consulting", Amount = 1234.5678 },
                new FinancialRecord { Description = "Software License", Amount = 9876.5432 },
                new FinancialRecord { Description = "Support", Amount = 250.125 }
            }
        };

        // 2. Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Header.
        builder.Writeln("Financial Report");
        builder.Writeln("-----------------");

        // Table header (outside the foreach loop).
        var headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Description");
        builder.InsertCell();
        builder.Writeln("Amount (rounded to 2 decimals)");
        builder.EndRow();
        builder.EndTable();

        // Data rows using LINQ Reporting tags.
        // The foreach block must enclose the entire table that repeats for each item.
        builder.Writeln("<<foreach [item in Items]>>");
        var dataTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[item.Description]>>");
        builder.InsertCell();
        // Use System.Math.Round static method inside the expression.
        builder.Writeln("<<[Math.Round(item.Amount, 2)]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // 3. Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // 4. Configure the ReportingEngine.
        var engine = new ReportingEngine();
        // Allow the engine to access static members of System.Math.
        engine.KnownTypes.Add(typeof(Math));

        // 5. Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
