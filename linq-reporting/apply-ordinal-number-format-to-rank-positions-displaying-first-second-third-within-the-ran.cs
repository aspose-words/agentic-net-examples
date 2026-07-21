using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data
        var players = new List<Player>
        {
            new Player { Name = "Alice", Score = 95 },
            new Player { Name = "Bob", Score = 87 },
            new Player { Name = "Charlie", Score = 78 },
            new Player { Name = "Diana", Score = 88 }
        };

        // Sort by score descending and assign ranks
        players.Sort((a, b) => b.Score.CompareTo(a.Score));
        for (int i = 0; i < players.Count; i++)
        {
            players[i].Rank = i + 1;
        }

        var model = new ReportModel { Rankings = players };

        // Create template document programmatically
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Ranking Report");
        builder.Writeln("");
        builder.Writeln("<<foreach [item in Rankings]>>");
        builder.Writeln("<<[item.Ordinal]>>: <<[item.Name]>> - Score: <<[item.Score]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template and build the report
        var template = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report
        var outputPath = "output.docx";
        template.Save(outputPath);
    }
}

public class ReportModel
{
    public List<Player> Rankings { get; set; } = new();
}

public class Player
{
    public string Name { get; set; } = "";
    public int Score { get; set; }
    public int Rank { get; set; }

    public string Ordinal => GetOrdinalString(Rank);

    private static string GetOrdinalString(int rank)
    {
        return rank switch
        {
            1 => "First",
            2 => "Second",
            3 => "Third",
            _ => $"{rank}th"
        };
    }
}
