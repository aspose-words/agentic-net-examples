using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Player
{
    public string Name { get; set; } = "";
    public int Score { get; set; }
    public int Rank { get; set; }

    // Returns ordinal representation of the rank (First, Second, Third, etc.).
    public string Ordinal => GetOrdinal(Rank);

    private static string GetOrdinal(int number)
    {
        return number switch
        {
            1 => "First",
            2 => "Second",
            3 => "Third",
            _ => $"{number}th"
        };
    }
}

public class Leaderboard
{
    public List<Player> Players { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for potential encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Step 1: Create the LINQ Reporting template programmatically.
        var templatePath = "LeaderboardTemplate.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Leaderboard");
        builder.Writeln("<<foreach [player in Players]>>");
        builder.Writeln("<<[player.Ordinal]>> - <<[player.Name]>> - Score: <<[player.Score]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Step 2: Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Step 3: Prepare sample data.
        var model = new Leaderboard
        {
            Players = new List<Player>
            {
                new Player { Name = "Alice",   Score = 95, Rank = 1 },
                new Player { Name = "Bob",     Score = 90, Rank = 2 },
                new Player { Name = "Charlie", Score = 85, Rank = 3 },
                new Player { Name = "David",   Score = 80, Rank = 4 }
            }
        };

        // Step 4: Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Step 5: Save the generated report.
        var outputPath = "LeaderboardReport.docx";
        reportDoc.Save(outputPath);
    }
}
