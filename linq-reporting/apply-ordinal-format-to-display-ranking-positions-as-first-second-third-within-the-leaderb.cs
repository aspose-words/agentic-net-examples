using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string templatePath = "LeaderboardTemplate.docx";
        string outputPath = "LeaderboardReport.docx";

        // -----------------------------------------------------------------
        // Step 1: Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Leaderboard");
        builder.Writeln("<<foreach [player in model.Leaderboard.Players]>>");
        builder.Writeln(
            "<<if [player.Rank == 1]>>First<</if>>" +
            "<<if [player.Rank == 2]>>Second<</if>>" +
            "<<if [player.Rank == 3]>>Third<</if>>" +
            "<<if [player.Rank > 3]>> <<[player.Rank]>> <</if>>. " +
            "<<[player.Name]>> - <<[player.Score]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Step 3: Prepare sample data.
        // -----------------------------------------------------------------
        var model = new Model
        {
            Leaderboard = new Leaderboard
            {
                Players = new List<Player>
                {
                    new Player { Name = "Alice", Score = 95 },
                    new Player { Name = "Bob", Score = 87 },
                    new Player { Name = "Charlie", Score = 78 },
                    new Player { Name = "Diana", Score = 65 },
                    new Player { Name = "Ethan", Score = 60 }
                }
            }
        };

        // Compute ranking based on descending scores.
        var ordered = model.Leaderboard.Players
            .OrderByDescending(p => p.Score)
            .Select((p, index) =>
            {
                p.Rank = index + 1;
                return p;
            })
            .ToList();

        model.Leaderboard.Players = ordered;

        // -----------------------------------------------------------------
        // Step 4: Build the report using LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // Step 5: Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model definitions.
// ---------------------------------------------------------------------
public class Model
{
    public Leaderboard Leaderboard { get; set; } = new();
}

public class Leaderboard
{
    public List<Player> Players { get; set; } = new();
}

public class Player
{
    public string Name { get; set; } = "";
    public int Score { get; set; }
    public int Rank { get; set; }
}
