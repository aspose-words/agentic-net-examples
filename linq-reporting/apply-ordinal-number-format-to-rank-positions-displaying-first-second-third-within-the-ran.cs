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
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output documents.
        string templatePath = "RankingTemplate.docx";
        string outputPath = "RankingReport.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Ranking Report");
        builder.Writeln();

        // Begin foreach loop over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        // Paragraph showing ordinal rank and name.
        builder.Writeln("<<[item.RankOrdinal]>>: <<[item.Name]>>");
        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new RankItem { Position = 1, Name = "Alice" },
                new RankItem { Position = 2, Name = "Bob" },
                new RankItem { Position = 3, Name = "Charlie" },
                new RankItem { Position = 4, Name = "Diana" }
            }
        };

        // Build the report using LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<RankItem> Items { get; set; } = new();
}

// Individual ranking item.
public class RankItem
{
    public int Position { get; set; }
    public string Name { get; set; } = string.Empty;

    // Returns ordinal word for the position (First, Second, Third, etc.).
    public string RankOrdinal => GetOrdinalWord(Position);

    private static string GetOrdinalWord(int position)
    {
        return position switch
        {
            1 => "First",
            2 => "Second",
            3 => "Third",
            _ => position + "th"
        };
    }
}
