using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains a histogram chart placeholder.
        Document doc = new Document("HistogramTemplate.docx");

        // Prepare a LINQ data source.
        // The template can reference the collection via the name "items".
        List<ReportItem> items = new List<ReportItem>
        {
            new ReportItem { Category = "Alpha",   Amount = 120 },
            new ReportItem { Category = "Beta",    Amount =  85 },
            new ReportItem { Category = "Gamma",   Amount = 150 },
            new ReportItem { Category = "Delta",   Amount =  95 },
            new ReportItem { Category = "Epsilon", Amount = 110 }
        };

        // Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, items, "items");

        // Locate the histogram chart that was inserted by the template.
        Shape chartShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Create the data array for the histogram from the LINQ query.
        double[] histogramValues = items.Select(i => (double)i.Amount).ToArray();

        // Add a series to the histogram chart.
        // For Histogram charts only X values are required; Y values are calculated automatically.
        chart.Series.Add("Amount", histogramValues);

        // Save the populated document.
        doc.Save("HistogramReport.docx");
    }

    // Simple data model used by the LINQ query.
    public class ReportItem
    {
        public string Category { get; set; }
        public int Amount { get; set; }
    }
}
