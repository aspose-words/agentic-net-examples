using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains a pie‑chart placeholder.
        Document doc = new Document("TemplatePieChart.docx");

        // Sample raw data.
        var sales = new List<Sale>
        {
            new Sale { Category = "Books",       Amount = 1200 },
            new Sale { Category = "Electronics",Amount = 3400 },
            new Sale { Category = "Clothing",   Amount = 2100 },
            new Sale { Category = "Books",       Amount = 800  },
            new Sale { Category = "Clothing",   Amount = 1500 }
        };

        // LINQ query that aggregates the data for the chart.
        var reportData = sales
            .GroupBy(s => s.Category)
            .Select(g => new
            {
                Category = g.Key,
                Amount   = g.Sum(x => x.Amount)
            })
            .ToList();

        // Populate the template with the LINQ result.
        // The third argument ("data") is the name used inside the template to reference the source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, reportData, "data");

        // Save the generated report.
        doc.Save("PieChartReport.docx");
    }

    // Simple POCO representing a sales record.
    public class Sale
    {
        public string Category { get; set; }
        public double Amount   { get; set; }
    }
}
