using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output.
        string templatePath = "LeaderboardTemplate.docx";
        string outputPath = "LeaderboardReport.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Leaderboard");
        builder.Writeln();

        // Table header.
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Rank");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Score");
        builder.EndRow();
        builder.EndTable();

        // Begin foreach over players.
        builder.Writeln("<<foreach [player in Leaderboard.Players]>>");

        // Table rows for each player.
        Table table = builder.StartTable();

        // Rank (ordinal).
        builder.InsertCell();
        builder.Writeln("<<[player.RankOrdinal]>>");

        // Name.
        builder.InsertCell();
        builder.Writeln("<<[player.Name]>>");

        // Score.
        builder.InsertCell();
        builder.Writeln("<<[player.Score]>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // Prepare sample data.
        Leaderboard model = new Leaderboard
        {
            Players = new List<Player>
            {
                new Player { Name = "Alice", Score = 95 },
                new Player { Name = "Bob", Score = 87 },
                new Player { Name = "Charlie", Score = 78 },
                new Player { Name = "Diana", Score = 65 }
            }
        };

        // Assign positions based on descending score.
        model.Players.Sort((a, b) => b.Score.CompareTo(a.Score));
        for (int i = 0; i < model.Players.Count; i++)
        {
            model.Players[i].Position = i + 1;
        }

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "Leaderboard");

        // Save the generated report.
        doc.Save(outputPath);
    }
}

// Root data model.
public class Leaderboard
{
    public List<Player> Players { get; set; } = new();
}

// Player data model.
public class Player
{
    public string Name { get; set; } = "";
    public int Score { get; set; }
    public int Position { get; set; }

    // Returns ordinal representation (First, Second, Third, etc.).
    public string RankOrdinal
    {
        get
        {
            return Position switch
            {
                1 => "First",
                2 => "Second",
                3 => "Third",
                _ => Position + GetOrdinalSuffix(Position)
            };
        }
    }

    private static string GetOrdinalSuffix(int number)
    {
        int absNumber = Math.Abs(number);
        int lastTwoDigits = absNumber % 100;
        if (lastTwoDigits >= 11 && lastTwoDigits <= 13)
            return "th";

        return (absNumber % 10) switch
        {
            1 => "st",
            2 => "nd",
            3 => "rd",
            _ => "th"
        };
    }
}
