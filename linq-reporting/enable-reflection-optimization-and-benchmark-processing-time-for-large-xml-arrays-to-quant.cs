using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a large XML data file.
        string xmlPath = Path.Combine(outputDir, "data.xml");
        GenerateLargeXml(xmlPath, 5000); // 5,000 items.

        // 2. Create a LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // 3. Benchmark without reflection optimization.
        ReportingEngine.UseReflectionOptimization = false;
        var (timeNoOpt, outputNoOpt) = BuildReport(templatePath, xmlPath, false, outputDir);

        // 4. Benchmark with reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;
        var (timeOpt, outputOpt) = BuildReport(templatePath, xmlPath, true, outputDir);

        // 5. Show results.
        Console.WriteLine($"Build time without reflection optimization: {timeNoOpt.TotalMilliseconds} ms");
        Console.WriteLine($"Build time with reflection optimization:    {timeOpt.TotalMilliseconds} ms");
        Console.WriteLine($"Outputs saved to: {outputNoOpt} and {outputOpt}");
    }

    // Generates an XML file with the specified number of <Order> elements.
    private static void GenerateLargeXml(string filePath, int count)
    {
        var root = new XElement("Orders");
        for (int i = 1; i <= count; i++)
        {
            var order = new XElement("Order",
                new XElement("Id", i),
                new XElement("CustomerName", $"Customer {i}")
            );
            root.Add(order);
        }
        var doc = new XDocument(root);
        doc.Save(filePath);
    }

    // Creates a simple Word template containing a foreach loop over Orders.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Builds the report, measures execution time, and saves the result.
    private static (TimeSpan elapsed, string outputPath) BuildReport(string templatePath, string xmlPath, bool useOpt, string outputDir)
    {
        // Load fresh template for each run.
        var doc = new Document(templatePath);

        // Load XML data source.
        var xmlDataSource = new XmlDataSource(xmlPath);

        var engine = new ReportingEngine();

        var stopwatch = Stopwatch.StartNew();
        // Root name matches the root element in the XML file.
        engine.BuildReport(doc, xmlDataSource, "Orders");
        stopwatch.Stop();

        string suffix = useOpt ? "opt" : "no_opt";
        string outputPath = Path.Combine(outputDir, $"report_{suffix}.docx");
        doc.Save(outputPath);

        return (stopwatch.Elapsed, outputPath);
    }
}
