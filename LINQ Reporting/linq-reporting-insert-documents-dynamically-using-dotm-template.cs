using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDynamicInsert
{
    // Simple data model that will be used as the data source for the DOTM template.
    public class ReportData
    {
        // The template will iterate over this collection and insert each document.
        public List<Document> Documents { get; set; }
    }

    public static class ReportGenerator
    {
        /// <summary>
        /// Generates a report by loading a DOTM template, inserting a collection of documents
        /// dynamically using the LINQ Reporting Engine, and saving the result.
        /// </summary>
        /// <param name="templatePath">Full path to the DOTM template file.</param>
        /// <param name="outputPath">Full path where the generated report will be saved.</param>
        /// <param name="documentPaths">Paths to the documents that should be inserted.</param>
        public static void Generate(string templatePath, string outputPath, IEnumerable<string> documentPaths)
        {
            // Load the DOTM template.
            Document template = new Document(templatePath);

            // Prepare the data source: load each document that will be inserted.
            // The ReportingEngine can work with a list of Document objects.
            var data = new ReportData
            {
                Documents = documentPaths
                    .Where(File.Exists)                     // Ensure the file exists.
                    .Select(path => new Document(path))     // Load each document.
                    .ToList()
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report.
            // The third parameter ("ds") is the name used in the template to reference the data source.
            // In the template you would have something like:
            //   <<foreach [in ds.Documents]>>
            //       <<doc [src.Document]>>
            //   <<end>>
            engine.BuildReport(template, data, "ds");

            // Save the generated report.
            template.Save(outputPath);
        }

        // Example usage.
        public static void Main()
        {
            // Path to the DOTM template that contains the reporting tags.
            string templatePath = @"C:\Templates\ReportTemplate.dotm";

            // Destination for the final report.
            string outputPath = @"C:\Reports\GeneratedReport.docx";

            // List of documents to be inserted into the report.
            var docsToInsert = new List<string>
            {
                @"C:\Docs\Section1.docx",
                @"C:\Docs\Section2.docx",
                @"C:\Docs\Section3.docx"
            };

            Generate(templatePath, outputPath, docsToInsert);

            Console.WriteLine("Report generated successfully.");
        }
    }
}
