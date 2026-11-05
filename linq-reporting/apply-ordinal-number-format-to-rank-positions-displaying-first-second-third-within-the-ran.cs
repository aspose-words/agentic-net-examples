using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the template document.
        const string templatePath = "RankingTemplate.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Ranking:");
        builder.Writeln("<<foreach [player in Players]>>");
        builder.Writeln("<<[player.Ordinal]>>. <<[player.Name]>> - <<[player.Score]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Players = GetSamplePlayers()
                .OrderByDescending(p => p.Score)
                .Select((p, index) =>
                {
                    p.Rank = index + 1;
                    return p;
                })
                .ToList()
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "RankingReport.docx";
        doc.Save(outputPath);
    }

    private static List<Player> GetSamplePlayers()
    {
        return new()
        {
            new Player { Name = "Alice",   Score = 95 },
            new Player { Name = "Bob",     Score = 87 },
            new Player { Name = "Charlie", Score = 92 },
            new Player { Name = "Diana",   Score = 78 }
        };
    }
}

// Root data model.
public class ReportModel
{
    public List<Player> Players { get; set; } = new();
}

// Data item model.
public class Player
{
    public string Name { get; set; } = "";
    public int Score { get; set; }
    public int Rank { get; set; }

    // Returns ordinal word based on Rank (First, Second, Third, etc.).
    public string Ordinal => GetOrdinalWord(Rank);

    private static string GetOrdinalWord(int rank) => rank switch
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
        _ => $"{rank}th"
    };
}
