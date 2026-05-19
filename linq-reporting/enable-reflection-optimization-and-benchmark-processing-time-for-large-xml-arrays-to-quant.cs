using System;
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
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and data files.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string xmlDataPath = Path.Combine(outputDir, "data.xml");
        string reportWithoutOptPath = Path.Combine(outputDir, "report_without_opt.docx");
        string reportWithOptPath = Path.Combine(outputDir, "report_with_opt.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Template uses a foreach loop over the "persons" collection.
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Generate a large XML data source (e.g., 10,000 person records).
        // -----------------------------------------------------------------
        const int recordCount = 10000;
        XDocument xmlDoc = new XDocument(
            new XElement("persons",
                GeneratePersonElements(recordCount)
            )
        );
        xmlDoc.Save(xmlDataPath, SaveOptions.DisableFormatting);

        // -----------------------------------------------------------------
        // 3. Benchmark without reflection optimization.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = false;
        Document docWithoutOpt = new Document(templatePath);
        XmlDataSource dataSourceWithoutOpt = new XmlDataSource(xmlDataPath);
        ReportingEngine engineWithoutOpt = new ReportingEngine();

        Stopwatch sw = Stopwatch.StartNew();
        engineWithoutOpt.BuildReport(docWithoutOpt, dataSourceWithoutOpt, "persons");
        sw.Stop();
        long timeWithoutOpt = sw.ElapsedMilliseconds;

        docWithoutOpt.Save(reportWithoutOptPath);

        // -----------------------------------------------------------------
        // 4. Benchmark with reflection optimization enabled.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = true;
        Document docWithOpt = new Document(templatePath);
        XmlDataSource dataSourceWithOpt = new XmlDataSource(xmlDataPath);
        ReportingEngine engineWithOpt = new ReportingEngine();

        sw.Restart();
        engineWithOpt.BuildReport(docWithOpt, dataSourceWithOpt, "persons");
        sw.Stop();
        long timeWithOpt = sw.ElapsedMilliseconds;

        docWithOpt.Save(reportWithOptPath);

        // -----------------------------------------------------------------
        // 5. Output the benchmark results.
        // -----------------------------------------------------------------
        Console.WriteLine($"Processing time without reflection optimization: {timeWithoutOpt} ms");
        Console.WriteLine($"Processing time with reflection optimization:    {timeWithOpt} ms");
    }

    // Helper method to generate a sequence of <person> elements.
    private static IEnumerable<XElement> GeneratePersonElements(int count)
    {
        for (int i = 1; i <= count; i++)
        {
            yield return new XElement("person",
                new XElement("Name", $"Person {i}"),
                new XElement("Age", (20 + (i % 50)).ToString())
            );
        }
    }
}
