using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Xml.Linq;

namespace ReflectionOptimizationBenchmark
{
    // Minimal placeholder for a document.
    class Document
    {
        public StringBuilder Content { get; } = new StringBuilder();

        public void Save(string path)
        {
            // Ensure directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, Content.ToString());
        }

        public Document Clone()
        {
            var clone = new Document();
            clone.Content.Append(this.Content);
            return clone;
        }
    }

    // Minimal placeholder for a data source based on XML.
    class XmlDataSource
    {
        public XDocument Document { get; }

        public XmlDataSource(string filePath)
        {
            Document = XDocument.Load(filePath);
        }

        public XElement Root => Document.Root!;
    }

    // Minimal reporting engine that simulates work.
    class ReportingEngine
    {
        public static bool UseReflectionOptimization { get; set; }

        public void BuildReport(Document doc, XmlDataSource dataSource, string rootElementName)
        {
            // Simulate different processing based on the optimization flag.
            var persons = dataSource.Root.Elements("person");
            foreach (var person in persons)
            {
                // Simulated work: read values and append to document content.
                var name = (string?)person.Element("Name") ?? string.Empty;
                var age = (string?)person.Element("Age") ?? string.Empty;

                if (UseReflectionOptimization)
                {
                    // Faster path (simulated).
                    doc.Content.AppendLine($"{name} ({age})");
                }
                else
                {
                    // Slower path (simulated extra work).
                    var combined = $"{name} ({age})";
                    // Simulate extra processing delay.
                    for (int i = 0; i < 3; i++)
                    {
                        combined = combined.Replace(" ", " ");
                    }
                    doc.Content.AppendLine(combined);
                }
            }
        }
    }

    class Program
    {
        // Use a directory relative to the current working directory.
        private static readonly string ArtifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");

        static void Main()
        {
            // Ensure the artifacts directory exists.
            Directory.CreateDirectory(ArtifactsDir);

            // 1. Prepare a large XML data file.
            string xmlPath = Path.Combine(ArtifactsDir, "LargeData.xml");
            GenerateLargeXml(xmlPath, elementCount: 200_000); // Adjust count as needed.

            // 2. Create a reporting template document in memory.
            Document template = CreateTemplateDocument();

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
        /// Creates a minimal placeholder document.
        /// </summary>
        private static Document CreateTemplateDocument()
        {
            // In a real scenario this would contain a template.
            // For this benchmark we just return an empty document.
            return new Document();
        }

        /// <summary>
        /// Builds the report using the provided template and data source,
        /// measuring the elapsed time of the BuildReport operation.
        /// </summary>
        private static TimeSpan BuildReportAndMeasure(Document template, XmlDataSource dataSource, string rootElementName)
        {
            // Clone the template to avoid modifying the original instance between runs.
            Document doc = template.Clone();

            var engine = new ReportingEngine();

            var stopwatch = Stopwatch.StartNew();

            // BuildReport populates the document with data from the XML source.
            engine.BuildReport(doc, dataSource, rootElementName);

            stopwatch.Stop();

            // Save the generated report to verify correctness.
            string outputPath = Path.Combine(ArtifactsDir,
                $"Report_{DateTime.Now:yyyyMMdd_HHmmss}_{(ReportingEngine.UseReflectionOptimization ? "Opt" : "NoOpt")}.txt");
            doc.Save(outputPath);

            return stopwatch.Elapsed;
        }
    }
}
