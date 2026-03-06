using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that already contains a bubble chart with reporting tags.
        Document doc = new Document("BubbleChartTemplate.docx");

        // Sample source collection – in a real scenario this could come from a database, XML, etc.
        var products = new List<Product>
        {
            new Product { Name = "Alpha",   Sales = 120, Profit = 30, MarketShare = 5 },
            new Product { Name = "Beta",    Sales = 200, Profit = 80, MarketShare = 12 },
            new Product { Name = "Gamma",   Sales = 150, Profit = 45, MarketShare = 8 }
        };

        // Use LINQ to transform the source data into the arrays required by a bubble chart.
        var bubbleSeries = new BubbleSeries
        {
            XValues = products.Select(p => (double)p.Sales).ToArray(),
            YValues = products.Select(p => (double)p.Profit).ToArray(),
            Sizes   = products.Select(p => (double)p.MarketShare).ToArray(),
            Labels  = products.Select(p => p.Name).ToArray()
        };

        // Wrap the series in an object that will be used as the data source for the report.
        var dataSource = new { Series = bubbleSeries };

        // Populate the template. The template should reference the data source as <<[ds.Series.XValues]>>, etc.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "ds");

        // Save the generated report.
        doc.Save("BubbleChartReport.docx");
    }

    // Simple POCO representing a row in the original collection.
    class Product
    {
        public string Name { get; set; }
        public int Sales { get; set; }
        public int Profit { get; set; }
        public int MarketShare { get; set; }
    }

    // Helper class matching the bubble‑chart series schema.
    class BubbleSeries
    {
        public double[] XValues { get; set; }
        public double[] YValues { get; set; }
        public double[] Sizes   { get; set; }
        public string[] Labels  { get; set; }
    }
}
