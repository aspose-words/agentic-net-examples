using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReflectionOptimizationBenchmark
{
    class Program
    {
        // Adjust these paths to point to your actual directories.
        private static readonly string MyDir = @"C:\Docs\Templates\";
        private static readonly string ArtifactsDir = @"C:\Docs\Artifacts\";

        static void Main()
        {
            // Ensure the artifacts directory exists.
            Directory.CreateDirectory(ArtifactsDir);

            // 1. Prepare a large XML data file.
            string xmlPath = Path.Combine(ArtifactsDir, "LargeData.xml");
            GenerateLargeXml(xmlPath, elementCount: 500_000); // Adjust count as needed.

            // 2. Load a reporting template document.
            // The template should contain a repeating region like: {{persons}}
            // where "persons" matches the root element name in the XML.
            string templatePath = Path.Combine(MyDir, "ReportingTemplate.docx");
            Document template = new Document(templatePath);

            // 3. Create the XML data source.
            XmlDataSource dataSource = new XmlDataSource(xmlPath);

            // 4. Benchmark with reflection optimization enabled.
            ReportingEngine.UseReflectionOptimization = true;
            TimeSpan timeWithOptimization = BuildReportAndMeasure(template, dataSource, "persons");

            // 5. Benchmark with reflection optimization disabled.
            ReportingEngine.UseReflectionOptimization = false;
            TimeSpan timeWithoutOptimization = BuildReportAndMeasure(template, dataSource, "persons");

            // 6. Output the results.
            Console.WriteLine($"Reflection optimization ENABLED : {timeWithOptimization.TotalMilliseconds} ms");
            Console.WriteLine($"Reflection optimization DISABLED: {timeWithoutOptimization.TotalMilliseconds} ms");
        }

        /// <summary>
        /// Generates a simple XML file containing a large number of repeated elements.
        /// Example structure:
        /// <persons>
        ///   <person>
        ///     <Name>John Doe 0</Name>
        ///     <Age>30</Age>
        ///   </person>
        ///   ...
        /// </persons>
        /// </summary>
        private static void GenerateLargeXml(string filePath, int elementCount)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<persons>");
            for (int i = 0; i < elementCount; i++)
            {
                sb.AppendLine("  <person>");
                sb.AppendLine($"    <Name>John Doe {i}</Name>");
                sb.AppendLine($"    <Age>{20 + (i % 50)}</Age>");
                sb.AppendLine("  </person>");
            }
            sb.AppendLine("</persons>");

            File.WriteAllText(filePath, sb.ToString());
        }

        /// <summary>
        /// Builds the report using the provided template and data source,
        /// measuring the elapsed time of the BuildReport operation.
        /// </summary>
        private static TimeSpan BuildReportAndMeasure(Document template, XmlDataSource dataSource, string rootElementName)
        {
            // Clone the template to avoid modifying the original instance between runs.
            Document doc = (Document)template.Clone();

            var engine = new ReportingEngine();

            var stopwatch = Stopwatch.StartNew();

            // BuildReport populates the document with data from the XML source.
            engine.BuildReport(doc, dataSource, rootElementName);

            stopwatch.Stop();

            // Optionally, save the generated report to verify correctness.
            // The file name includes a timestamp to avoid overwriting.
            string outputPath = Path.Combine(ArtifactsDir,
                $"Report_{DateTime.Now:yyyyMMdd_HHmmss}_{(ReportingEngine.UseReflectionOptimization ? "Opt" : "NoOpt")}.docx");
            doc.Save(outputPath);

            return stopwatch.Elapsed;
        }
    }
}
