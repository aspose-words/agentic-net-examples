using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Player
{
    public string Name { get; set; } = "";
    public int Score { get; set; }
}

public class ReportModel
{
    public List<Player> Players { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // ---------- Create the template ----------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Heading
        builder.Writeln("Ranking:");

        // Define a list whose numbering style is ordinal text (First, Second, Third, …)
        List list = template.Lists.Add(ListTemplate.NumberDefault);
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.OrdinalText; // First, Second, Third, …
        level.StartAt = 1;

        // Begin the LINQ Reporting foreach block
        builder.Writeln("<<foreach [player in Players]>>");

        // Apply the list to the paragraph that will be repeated for each player
        builder.ListFormat.List = list;
        builder.Writeln("<<[player.Name]>> - <<[player.Score]>>");
        builder.ListFormat.RemoveNumbers(); // Stop list formatting after the block

        // End the foreach block
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        const string templatePath = "RankingTemplate.docx";
        template.Save(templatePath);

        // ---------- Prepare the data ----------
        ReportModel model = new ReportModel
        {
            Players = new List<Player>
            {
                new Player { Name = "Alice",   Score = 95 },
                new Player { Name = "Bob",     Score = 88 },
                new Player { Name = "Charlie", Score = 82 }
            }
        };

        // ---------- Build the report ----------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the final report
        const string outputPath = "RankingReport.docx";
        report.Save(outputPath);

        // Inform the user (no input required)
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
