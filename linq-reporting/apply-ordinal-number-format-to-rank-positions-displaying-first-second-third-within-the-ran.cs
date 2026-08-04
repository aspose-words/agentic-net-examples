using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Player
{
    public int Rank { get; set; }
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Player> Players { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for legacy encodings (required by Aspose.Words in some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Players = new List<Player>
            {
                new Player { Rank = 1, Name = "Alice" },
                new Player { Rank = 2, Name = "Bob" },
                new Player { Rank = 3, Name = "Charlie" },
                new Player { Rank = 4, Name = "Diana" }
            }
        };

        // Create a template document programmatically
        string templatePath = "RankingTemplate.docx";
        CreateTemplate(templatePath);

        // Load the template
        Document doc = new Document(templatePath);

        // Build the report using LINQ Reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save("RankingReport.docx");
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Ranking Report");
        builder.Writeln();

        // Begin foreach loop over Players collection
        builder.Writeln("<<foreach [player in Players]>>");

        // Use ordinal text format for the rank (First, Second, Third, ...)
        builder.Writeln("<<[player.Rank]:ordinalText>>. <<[player.Name]>>");

        // End foreach loop
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(filePath);
    }
}
