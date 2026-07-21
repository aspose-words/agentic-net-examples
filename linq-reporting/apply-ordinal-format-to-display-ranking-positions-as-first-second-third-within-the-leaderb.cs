using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new Leaderboard
        {
            Players = new List<Player>
            {
                new Player { Name = "Alice", Score = 150 },
                new Player { Name = "Bob",   Score = 120 },
                new Player { Name = "Carol", Score = 100 }
            }
        };

        // Assign ranking positions (1‑based).
        for (int i = 0; i < model.Players.Count; i++)
            model.Players[i].Rank = i + 1;

        // Create the LINQ Reporting template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "LeaderboardTemplate.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Leaderboard");
        builder.Writeln();

        // Start a foreach block over the Players collection.
        builder.Writeln("<<foreach [player in Players]>>");
        // Use the ordinalText format to display 1 → First, 2 → Second, etc.
        builder.Writeln("Rank: <<[player.Rank]:ordinalText>>  Name: <<[player.Name]>>  Score: <<[player.Score]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "LeaderboardReport.docx");
        reportDoc.Save(reportPath);
    }
}

// Root data model.
public class Leaderboard
{
    public List<Player> Players { get; set; } = new();
}

// Individual player entry.
public class Player
{
    public int Rank { get; set; }          // Ranking position (numeric).
    public string Name { get; set; } = ""; // Player name.
    public int Score { get; set; }         // Player score.
}
