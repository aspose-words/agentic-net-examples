using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for the Table class

namespace AsposeWordsLinqReporting
{
    // Data model for a player in the leaderboard.
    public class Player
    {
        public string Name { get; set; } = string.Empty;
        public int Score { get; set; }
        public int Rank { get; set; }
    }

    // Wrapper class that will be passed as the root data source.
    public class Leaderboard
    {
        public List<Player> Players { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var leaderboard = new Leaderboard
            {
                Players = new List<Player>
                {
                    new Player { Name = "Alice",   Score = 95 },
                    new Player { Name = "Bob",     Score = 87 },
                    new Player { Name = "Charlie", Score = 78 }
                }
            };

            // Assign ranking positions based on the order in the list (1‑based rank).
            for (int i = 0; i < leaderboard.Players.Count; i++)
                leaderboard.Players[i].Rank = i + 1;

            // Create a template document programmatically.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Leaderboard");
            builder.Writeln("<<foreach [player in Model.Players]>>");

            // Table header.
            Table table = builder.StartTable();
            builder.InsertCell(); builder.Writeln("Rank");
            builder.InsertCell(); builder.Writeln("Name");
            builder.InsertCell(); builder.Writeln("Score");
            builder.EndRow();

            // Table row with ordinal formatting for the rank.
            builder.InsertCell(); builder.Writeln("<<[player.Rank]:ordinalText>>");
            builder.InsertCell(); builder.Writeln("<<[player.Name]>>");
            builder.InsertCell(); builder.Writeln("<<[player.Score]>>");
            builder.EndRow();

            builder.EndTable();

            builder.Writeln("<</foreach>>");

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(doc, leaderboard, "Model");

            // Save the generated report.
            doc.Save("LeaderboardReport.docx");
        }
    }
}
