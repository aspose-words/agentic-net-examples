using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing.Charts;

public class RegionPopulation
{
    public string Level1 { get; set; }
    public string Level2 { get; set; }
    public string Level3 { get; set; }
    public double Value { get; set; }
}

// Item that matches the structure expected by the template (multilevel value + numeric value)
public class ReportItem
{
    public ChartMultilevelValue Category { get; set; }
    public double Amount { get; set; }
}

// Root data source passed to the ReportingEngine
public class ReportModel
{
    public List<ReportItem> Data { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Load the DOCX template that contains a Treemap chart with reporting tags.
        Document template = new Document("TreemapTemplate.docx");

        // -----------------------------------------------------------------
        // 1. Prepare raw hierarchical data.
        // -----------------------------------------------------------------
        List<RegionPopulation> rawData = new List<RegionPopulation>
        {
            new RegionPopulation { Level1 = "Asia",            Level2 = "China",          Value = 1409670000 },
            new RegionPopulation { Level1 = "Asia",            Level2 = "India",          Value = 1400744000 },
            new RegionPopulation { Level1 = "Asia",            Level2 = "Indonesia",      Value = 279118866 },
            new RegionPopulation { Level1 = "Africa",          Level2 = "Nigeria",        Value = 223800000 },
            new RegionPopulation { Level1 = "Europe",          Level2 = "Germany",        Value = 84607016 },
            new RegionPopulation { Level1 = "Northern America",Level2 = "United States",  Level3 = "Other", Value = 335893238 },
            new RegionPopulation { Level1 = "Oceania",         Level2 = null,            Value = 42000000 }
        };

        // -----------------------------------------------------------------
        // 2. Shape the data with LINQ into the format required by the template.
        // -----------------------------------------------------------------
        List<ReportItem> reportItems = rawData
            .Select(r => new ReportItem
            {
                // ChartMultilevelValue can accept 1‑3 levels; missing levels are supplied as empty strings.
                Category = new ChartMultilevelValue(
                    r.Level1 ?? string.Empty,
                    r.Level2 ?? string.Empty,
                    r.Level3 ?? string.Empty),
                Amount = r.Value
            })
            .ToList();

        // -----------------------------------------------------------------
        // 3. Create the model that will be passed to the ReportingEngine.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel { Data = reportItems };

        // -----------------------------------------------------------------
        // 4. Build the report. The template refers to the data source by the name "model".
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the populated document.
        // -----------------------------------------------------------------
        template.Save("TreemapReport.docx");
    }
}
