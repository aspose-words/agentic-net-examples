using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data.
        const string xmlFileName = "orders.xml";
        CreateSampleXml(xmlFileName, 500); // 500 orders for demonstration.

        // Load XML and build data model while reporting progress.
        ReportModel model = new();
        LoadXmlWithProgress(xmlFileName, model, ReportProgress);

        // Set additional model data required by the template.
        model.CurrentDateTime = DateTime.Now;

        // Create a Word template with LINQ Reporting tags.
        const string templateFileName = "template.docx";
        CreateTemplate(templateFileName);

        // Load the template document.
        Document doc = new(templateFileName);

        // Build the report.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.InlineErrorMessages; // optional, provides detailed errors
        ReportProgress(0); // Start of report generation.
        bool success = engine.BuildReport(doc, model, "model");
        ReportProgress(100); // End of report generation.

        // Save the generated report.
        const string outputFileName = "report.docx";
        doc.Save(outputFileName, SaveFormat.Docx);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputFileName}'.");
    }

    // Progress callback that writes percentage to the console.
    private static void ReportProgress(int percent)
    {
        Console.WriteLine($"Progress: {percent}%");
    }

    // Creates a sample XML file with a specified number of orders.
    private static void CreateSampleXml(string filePath, int orderCount)
    {
        XElement root = new("Orders");
        for (int i = 1; i <= orderCount; i++)
        {
            root.Add(new XElement("Order",
                new XElement("Id", i),
                new XElement("CustomerName", $"Customer {i}")
            ));
        }

        XDocument doc = new(root);
        doc.Save(filePath);
    }

    // Loads XML data into the model while invoking the progress callback.
    private static void LoadXmlWithProgress(string filePath, ReportModel model, Action<int> progressCallback)
    {
        XDocument xdoc = XDocument.Load(filePath);
        var orderElements = xdoc.Root?.Elements("Order") ?? Enumerable.Empty<XElement>();
        int total = orderElements.Count();
        int processed = 0;

        foreach (var elem in orderElements)
        {
            Order order = new()
            {
                Id = (int?)elem.Element("Id") ?? 0,
                CustomerName = (string?)elem.Element("CustomerName") ?? string.Empty
            };
            model.Orders.Add(order);
            processed++;
            int percent = (int)((double)processed / total * 100);
            progressCallback(percent);
        }
    }

    // Creates a Word document template with LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        builder.Writeln("Orders Report");
        builder.Writeln("Generated on: <<[model.CurrentDateTime]>>");
        builder.Writeln();

        // Begin foreach loop over Orders collection.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath, SaveFormat.Docx);
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
    public DateTime CurrentDateTime { get; set; } = DateTime.Now;
}

// Individual order data.
public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = string.Empty;
}
