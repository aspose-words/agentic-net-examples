using System;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportGeneration
{
    public class AsyncReportBuilder
    {
        /// <summary>
        /// Generates a report from a template document without blocking the calling thread.
        /// </summary>
        /// <param name="templatePath">Full path to the template .docx file.</param>
        /// <param name="dataSource">The data source object used by the reporting engine.</param>
        /// <param name="outputPath">Full path where the generated report will be saved.</param>
        /// <returns>A task that completes when the report has been built and saved.</returns>
        public static async Task BuildReportAsync(string templatePath, object dataSource, string outputPath)
        {
            // Load the template document.
            Document template = new Document(templatePath);

            // Create a reporting engine instance.
            ReportingEngine engine = new ReportingEngine();

            // Optionally configure engine options (e.g., remove empty paragraphs).
            // engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Run the potentially long‑running BuildReport operation on a background thread.
            bool success = await Task.Run(() => engine.BuildReport(template, dataSource));

            // If the build succeeded (or you want to ignore the flag), save the result.
            // The Save method determines the format from the file extension.
            template.Save(outputPath);
        }

        // Example usage.
        public static async Task Example()
        {
            string templateFile = @"C:\Templates\ReportTemplate.docx";
            string outputFile = @"C:\Reports\GeneratedReport.docx";

            // Example data source – replace with your actual data.
            var data = new
            {
                Title = "Quarterly Sales",
                Date = DateTime.Now,
                Sales = new[]
                {
                    new { Region = "North", Amount = 12345.67 },
                    new { Region = "South", Amount = 8901.23 }
                }
            };

            await BuildReportAsync(templateFile, data, outputFile);
        }
    }

    public class Program
    {
        // C# 7.1+ supports an async Main returning Task.
        public static async Task Main(string[] args)
        {
            await AsyncReportBuilder.Example();
        }
    }
}
