using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class ReportData
{
    public string Title { get; set; }
    public Item[] Items { get; set; }
}

public class Item
{
    public string Name { get; set; }
    public int Quantity { get; set; }
    public double Price { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Load the DOCX template that contains LINQ tags.
        Document doc = new Document("Template.docx");

        // Prepare a data source that matches the tags used in the template.
        var data = new ReportData
        {
            Title = "Quarterly Report",
            Items = new[]
            {
                new Item { Name = "Product A", Quantity = 120, Price = 9.99 },
                new Item { Name = "Product B", Quantity = 85,  Price = 14.50 }
            }
        };

        // Populate the template with the data using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The third argument ("ds") is the name by which the template references the data source.
        engine.BuildReport(doc, data, "ds");

        // Save the resulting document as a PDF file.
        doc.Save("Report.pdf", SaveFormat.Pdf);
    }
}
