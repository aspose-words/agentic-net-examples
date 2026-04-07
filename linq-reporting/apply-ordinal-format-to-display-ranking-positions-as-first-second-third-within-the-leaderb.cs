using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "LeaderboardTemplate.docx";
        const string reportPath = "LeaderboardReport.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Leaderboard:");

        // Begin the foreach loop over the Players collection.
        builder.Writeln("<<foreach [player in Players]>>");

        // Each line will display the ordinal rank, name and score.
        // Example output after the report is built:
        // First: Alice - 1500
        // Second: Bob - 1200
        // Third: Carol - 900
        builder.Writeln("<<[player.OrdinalRank]>>: <<[player.Name]>> - <<[player.Score]>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model.
        // -----------------------------------------------------------------
        var model = new Leaderboard
        {
            Players = new List<Player>
            {
                new Player { Name = "Alice", Score = 1500, Rank = 1 },
                new Player { Name = "Bob",   Score = 1200, Rank = 2 },
                new Player { Name = "Carol", Score =  900, Rank = 3 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class Leaderboard
{
    // Collection of players that will be iterated in the template.
    public List<Player> Players { get; set; } = new();
}

public class Player
{
    // Player's display name.
    public string Name { get; set; } = string.Empty;

    // Player's score.
    public int Score { get; set; }

    // Numeric rank (1‑based).
    public int Rank { get; set; }

    // Ordinal representation of the rank (First, Second, Third, ...).
    public string OrdinalRank => GetOrdinal(Rank);

    // Converts a number to its ordinal text representation.
    private static string GetOrdinal(int number)
    {
        return number switch
        {
            1 => "First",
            2 => "Second",
            3 => "Third",
            4 => "Fourth",
            5 => "Fifth",
            6 => "Sixth",
            7 => "Seventh",
            8 => "Eighth",
            9 => "Ninth",
            10 => "Tenth",
            _ => number + "th"
        };
    }
}
