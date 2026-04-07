using System;
using System.Diagnostics;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for template, data and reports.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string xmlPath = Path.Combine(outputDir, "Data.xml");
        string reportOptimizedPath = Path.Combine(outputDir, "Report_Optimized.docx");
        string reportNonOptimizedPath = Path.Combine(outputDir, "Report_NonOptimized.docx");

        // 1. Create a simple LINQ Reporting template.
        CreateTemplate(templatePath);

        // 2. Generate a large XML data source.
        const int recordCount = 5000; // Adjust for desired size.
        GenerateLargeXml(xmlPath, recordCount);

        // 3. Benchmark with reflection optimization enabled.
        ReportingEngine.UseReflectionOptimization = true;
        TimeSpan optimizedTime = BuildReport(templatePath, xmlPath, "Orders", reportOptimizedPath);

        // 4. Benchmark with reflection optimization disabled.
        ReportingEngine.UseReflectionOptimization = false;
        TimeSpan nonOptimizedTime = BuildReport(templatePath, xmlPath, "Orders", reportNonOptimizedPath);

        // 5. Output the measured times.
        Console.WriteLine($"Reflection optimization enabled : {optimizedTime.TotalMilliseconds} ms");
        Console.WriteLine($"Reflection optimization disabled: {nonOptimizedTime.TotalMilliseconds} ms");
    }

    // Creates a Word document containing a foreach tag that iterates over <Order> elements.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Write a heading.
        builder.Writeln("Order Report");
        builder.Writeln();

        // LINQ Reporting foreach tag.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Id: <<[order.Id]>> , Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Generates an XML file with a large number of <Order> elements.
    private static void GenerateLargeXml(string filePath, int count)
    {
        XElement root = new("Orders");
        for (int i = 1; i <= count; i++)
        {
            XElement order = new("Order",
                new XElement("Id", i),
                new XElement("Amount", (i * 10).ToString()));
            root.Add(order);
        }

        XDocument doc = new(root);
        doc.Save(filePath);
    }

    // Loads the template, builds the report, saves it, and returns the elapsed time.
    private static TimeSpan BuildReport(string templatePath, string xmlPath, string rootName, string outputPath)
    {
        // Load the template document.
        Document template = new(templatePath);

        // Create the XML data source.
        XmlDataSource dataSource = new(xmlPath);

        // Prepare the reporting engine.
        ReportingEngine engine = new();

        // Measure the BuildReport execution time.
        Stopwatch sw = Stopwatch.StartNew();
        engine.BuildReport(template, dataSource, rootName);
        sw.Stop();

        // Save the generated report.
        template.Save(outputPath);

        return sw.Elapsed;
    }
}
