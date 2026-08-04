using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Leaderboard");

        // Begin foreach loop over Players collection.
        builder.Writeln("<<foreach [player in Players]>>");
        // Use ordinal text format for the Position property (First, Second, Third, ...).
        builder.Writeln("<<[player.Position]:ordinalText>>. <<[player.Name]>>");
        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to a temporary file.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "LeaderboardTemplate.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        Leaderboard model = new()
        {
            Players = new()
            {
                new Player { Position = 1, Name = "Alice" },
                new Player { Position = 2, Name = "Bob" },
                new Player { Position = 3, Name = "Charlie" }
            }
        };

        // Build the report using LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "LeaderboardReport.docx");
        doc.Save(outputPath);
    }
}

// Root data model for the report.
public class Leaderboard
{
    public List<Player> Players { get; set; } = new();
}

// Individual player entry.
public class Player
{
    public int Position { get; set; }
    public string Name { get; set; } = string.Empty;
}
