using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace Example
{
    public class ReportGenerator
    {
        /// <summary>
        /// Generates a separate report for each DataRow in the first table of the supplied DataSet.
        /// Each report is created from the same template document and saved to the specified output folder.
        /// </summary>
        /// <param name="dataSet">DataSet containing at least one DataTable with rows to process.</param>
        /// <param name="templatePath">Full path to the Word template file (e.g., .docx) that contains the reporting tags.</param>
        /// <param name="outputFolder">Folder where individual reports will be saved.</param>
        public void GenerateReports(DataSet dataSet, string templatePath, string outputFolder)
        {
            if (dataSet == null || dataSet.Tables.Count == 0)
                throw new ArgumentException("DataSet must contain at least one DataTable.", nameof(dataSet));

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load the template document once – it will be cloned for each row.
            Document template = new Document(templatePath);

            // Get the first table (you can change the index if needed).
            DataTable table = dataSet.Tables[0];

            // Iterate over each DataRow and build a separate report.
            for (int i = 0; i < table.Rows.Count; i++)
            {
                DataRow row = table.Rows[i];

                // Clone the template so that each iteration works with a fresh document.
                Document report = (Document)template.Clone(true);

                // Use ReportingEngine to populate the document with the current row.
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(report, row);

                // Construct a unique file name for the generated report.
                string outputPath = Path.Combine(outputFolder, $"Report_Row_{i + 1}.docx");

                // Save the populated document.
                report.Save(outputPath);
            }
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point required by the .NET compiler. Demonstrates usage of ReportGenerator.
        /// </summary>
        static void Main(string[] args)
        {
            // Prepare a simple DataSet with one DataTable for demonstration purposes.
            DataSet ds = new DataSet();
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Position", typeof(string));
            dt.Rows.Add("Alice", "Manager");
            dt.Rows.Add("Bob", "Developer");
            ds.Tables.Add(dt);

            // Paths – adjust to your environment.
            string templateFile = Path.Combine(Environment.CurrentDirectory, "ReportTemplate.docx");
            string outputDir = Path.Combine(Environment.CurrentDirectory, "GeneratedReports");

            // Ensure a template file exists; for a real run replace with an actual .docx containing reporting tags.
            if (!File.Exists(templateFile))
            {
                Console.WriteLine($"Template file not found: {templateFile}");
                return;
            }

            // Generate the reports.
            ReportGenerator generator = new ReportGenerator();
            generator.GenerateReports(ds, templateFile, outputDir);

            Console.WriteLine($"Reports generated in: {outputDir}");
        }
    }
}
