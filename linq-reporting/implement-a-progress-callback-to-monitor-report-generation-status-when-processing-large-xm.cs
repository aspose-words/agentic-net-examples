using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        string xmlPath = Path.Combine(workDir, "data.xml");
        string templatePath = Path.Combine(workDir, "template.docx");
        string reportPath = Path.Combine(workDir, "report.docx");

        // Create sample large XML data set
        CreateSampleXml(xmlPath, 200);

        // Load data with progress reporting
        Order order = LoadProducts(xmlPath, (processed, total) =>
        {
            Console.WriteLine($"Loading XML: {processed}/{total} ({processed * 100 / total}%)");
        });

        // Build template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Description: <<[p.Description]>>");
        builder.Writeln("<</foreach>>");
        template.Save(templatePath);

        // Load template (optional, we already have it)
        Document doc = new Document(templatePath);

        // Configure reporting engine
        ReportingEngine.UseReflectionOptimization = true;
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // Build report
        engine.BuildReport(doc, order, "order");

        // Save final report
        doc.Save(reportPath);
        Console.WriteLine($"Report generated at: {reportPath}");
    }

    private static void CreateSampleXml(string filePath, int count)
    {
        using (var writer = XmlWriter.Create(filePath, new XmlWriterSettings { Indent = true, Encoding = Encoding.UTF8 }))
        {
            writer.WriteStartDocument();
            writer.WriteStartElement("Products");
            for (int i = 1; i <= count; i++)
            {
                writer.WriteStartElement("Product");
                writer.WriteElementString("Name", $"Product {i}");
                writer.WriteElementString("Description", $"Description for product {i}");
                writer.WriteEndElement(); // Product
            }
            writer.WriteEndElement(); // Products
            writer.WriteEndDocument();
        }
    }

    private static Order LoadProducts(string xmlPath, Action<int, int> progressCallback)
    {
        int total = 0;
        using (var countReader = XmlReader.Create(xmlPath))
        {
            while (countReader.Read())
            {
                if (countReader.NodeType == XmlNodeType.Element && countReader.Name == "Product")
                    total++;
            }
        }

        var order = new Order();
        int processed = 0;
        using (var reader = XmlReader.Create(xmlPath))
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "Product")
                {
                    var product = new Product();
                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            switch (reader.Name)
                            {
                                case "Name":
                                    product.Name = reader.ReadElementContentAsString();
                                    break;
                                case "Description":
                                    product.Description = reader.ReadElementContentAsString();
                                    break;
                            }
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "Product")
                        {
                            break;
                        }
                    }
                    order.Products.Add(product);
                    processed++;
                    progressCallback?.Invoke(processed, total);
                }
            }
        }
        return order;
    }
}

public class Order
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
}
