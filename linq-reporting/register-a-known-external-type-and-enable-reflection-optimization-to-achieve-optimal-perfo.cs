using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Generate sample XML data
        string xmlPath = "Orders.xml";
        GenerateSampleXml(xmlPath, 5000);

        // 2. Load data into the model
        ReportModel model = LoadModelFromXml(xmlPath);

        // 3. Create a Word template programmatically
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // 4. Enable reflection optimization
        ReportingEngine.UseReflectionOptimization = true;
        ReportingEngine engine = new ReportingEngine();

        // 5. Load the template and build the report
        Document doc = new Document(templatePath);
        bool success = engine.BuildReport(doc, model, "model");

        // 6. Save the generated report
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        // 7. Indicate completion
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }

    private static void GenerateSampleXml(string filePath, int count)
    {
        var orders = new List<Order>();
        DateTime start = new DateTime(2023, 1, 1);
        for (int i = 1; i <= count; i++)
        {
            orders.Add(new Order
            {
                Id = i,
                Date = start.AddDays(i)
            });
        }

        var wrapper = new XmlOrdersWrapper { Orders = orders };
        var serializer = new XmlSerializer(typeof(XmlOrdersWrapper));
        using var writer = new StreamWriter(filePath);
        serializer.Serialize(writer, wrapper);
    }

    private static ReportModel LoadModelFromXml(string filePath)
    {
        var serializer = new XmlSerializer(typeof(XmlOrdersWrapper));
        using var reader = new StreamReader(filePath);
        var wrapper = (XmlOrdersWrapper)serializer.Deserialize(reader)!;
        return new ReportModel { Orders = wrapper.Orders };
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Orders Report");
        builder.Writeln("==============");
        builder.Writeln();

        // Begin foreach over Orders collection
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Date: <<[order.FormattedDate]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Wrapper used for XML serialization (matches the XML file structure)
[XmlRoot("Orders")]
public class XmlOrdersWrapper
{
    [XmlElement("Order")]
    public List<Order> Orders { get; set; } = new();
}

// Public data model used by the reporting engine
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Simple order class
public class Order
{
    public int Id { get; set; }
    public DateTime Date { get; set; }

    // Helper property for formatted date (used in the template)
    public string FormattedDate => Date.ToString("yyyy-MM-dd");
}
