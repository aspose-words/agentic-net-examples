using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class RankingItem
{
    public string Name { get; set; } = "";
    public int Position { get; set; }
}

public class ReportModel
{
    public List<RankingItem> Rankings { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Rankings = new List<RankingItem>
            {
                new RankingItem { Position = 1, Name = "Alice" },
                new RankingItem { Position = 2, Name = "Bob" },
                new RankingItem { Position = 3, Name = "Charlie" }
            }
        };

        // -----------------------------------------------------------------
        // Step 1: Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Ranking Report");
        builder.Writeln(); // Empty line for readability.

        // Begin foreach loop over Rankings collection.
        builder.Writeln("<<foreach [item in Rankings]>>");
        // Apply ordinal text format to the Position field (First, Second, Third, ...).
        builder.Writeln("<<[item.Position]:ordinalText>> - <<[item.Name]>>");
        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "RankingTemplate.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Default options.

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "RankingReport.docx";
        reportDoc.Save(outputPath);
    }
}
