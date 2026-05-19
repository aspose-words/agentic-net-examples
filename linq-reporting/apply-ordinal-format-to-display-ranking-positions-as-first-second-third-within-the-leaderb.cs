using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Player
{
    public int Rank { get; set; }
    public string Name { get; set; } = "";
    public int Score { get; set; }
}

public class LeaderboardModel
{
    public List<Player> Players { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for extended encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var model = new LeaderboardModel();
        var samplePlayers = new List<(string Name, int Score)>
        {
            ("Alice", 95),
            ("Bob", 87),
            ("Charlie", 92),
            ("Diana", 78)
        };

        // Sort by score descending and assign rank.
        int rank = 1;
        foreach (var p in samplePlayers.OrderByDescending(p => p.Score))
        {
            model.Players.Add(new Player { Rank = rank++, Name = p.Name, Score = p.Score });
        }

        // Build template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Leaderboard");
        builder.Writeln("<<foreach [player in Players]>>");
        // Display rank as ordinal text (First, Second, Third, ...).
        builder.Writeln("<<[player.Rank]:ordinalText>>. <<[player.Name]>> - <<[player.Score]>> points");
        builder.Writeln("<</foreach>>");

        // Generate the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("LeaderboardReport.docx");
    }
}
