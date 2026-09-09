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
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Ranking Report");
        builder.Writeln("<<foreach [p in Players]>>");
        // Apply ordinal text format (First, Second, Third, ...) to the Rank field.
        builder.Writeln("<<[p.Rank]:ordinalText>>. <<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back for report generation.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
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

        // -----------------------------------------------------------------
        // 4. Build the report using Aspose.Words LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable properties initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Player> Players { get; set; } = new();
}

public class Player
{
    public int Rank { get; set; }
    public string Name { get; set; } = string.Empty;
}
