using System;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsBenchmark
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the template document and the large JSON data file.
            string templatePath = @"C:\Data\ReportTemplate.docx";
            string jsonPath = @"C:\Data\LargeDataSet.json";

            // Run benchmark with reflection optimization enabled.
            BenchmarkReflectionOptimization(templatePath, jsonPath, true);

            // Run benchmark with reflection optimization disabled for comparison.
            BenchmarkReflectionOptimization(templatePath, jsonPath, false);
        }

        /// <summary>
        /// Benchmarks the time taken by ReportingEngine.BuildReport for a given JSON data source,
        /// optionally enabling or disabling reflection optimization.
        /// </summary>
        /// <param name="templatePath">Path to the Word template document.</param>
        /// <param name="jsonPath">Path to the large JSON file.</param>
        /// <param name="enableOptimization">If true, enables reflection optimization.</param>
        static void BenchmarkReflectionOptimization(string templatePath, string jsonPath, bool enableOptimization)
        {
            // Set the static property to control reflection optimization.
            ReportingEngine.UseReflectionOptimization = enableOptimization;

            // Load the template document (creation rule).
            Document doc = new Document(templatePath);

            // Load JSON data source from file (loading rule).
            JsonDataSource dataSource = new JsonDataSource(jsonPath);

            // Create a ReportingEngine instance.
            ReportingEngine engine = new ReportingEngine();

            // Measure the time taken to build the report.
            Stopwatch sw = Stopwatch.StartNew();

            // Build the report using the JSON data source.
            // The root object name ("root") should match the JSON structure.
            engine.BuildReport(doc, dataSource, "root");

            sw.Stop();

            // Output the elapsed time.
            Console.WriteLine($"Reflection optimization {(enableOptimization ? "enabled" : "disabled")}: {sw.ElapsedMilliseconds} ms");

            // Save the generated report (saving rule). Use a distinct file name for each run.
            string outputPath = enableOptimization
                ? @"C:\Data\Report_WithOptimization.docx"
                : @"C:\Data\Report_WithoutOptimization.docx";

            // Use default save options; no special options required.
            doc.Save(outputPath);
        }
    }
}
