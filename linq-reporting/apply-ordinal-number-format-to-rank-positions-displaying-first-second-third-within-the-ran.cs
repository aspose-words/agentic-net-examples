using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Player
{
    public int Position { get; set; }
    public string Name { get; set; } = string.Empty;

    // Returns ordinal text for the position (First, Second, Third, etc.).
    public string RankText
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
        int abs = Math.Abs(number);
        int lastTwo = abs % 100;
        if (lastTwo >= 11 && lastTwo <= 13)
            return "th";

        return (abs % 10) switch
        {
            1 => "st",
            2 => "nd",
            3 => "rd",
            _ => "th"
        };
    }
}

public class ReportModel
{
    public List<Player> Players { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create the template document with LINQ Reporting tags.
        var templatePath = "RankingTemplate.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Ranking Report");
        builder.Writeln("<<foreach [player in Players]>>");
        builder.Writeln("<<[player.RankText]>> - <<[player.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // 2. Load the template for report generation.
        var doc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new ReportModel
        {
            Players = new List<Player>
            {
                new Player { Position = 1, Name = "Alice" },
                new Player { Position = 2, Name = "Bob" },
                new Player { Position = 3, Name = "Charlie" },
                new Player { Position = 4, Name = "Diana" }
            }
        };

        // 4. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        doc.Save("RankingReport.docx");
    }
}
