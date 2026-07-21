using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting; // for CsvDataLoadOptions

namespace AsposeWordsLinqReportingBatchCsv
{
    // Model representing a CSV file.
    public class FileModel
    {
        public string FileName { get; set; } = string.Empty;
    }

    // Row representation matching the CSV columns.
    public class Row
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public double Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare folders.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);
            string dataDir = Path.Combine(workDir, "Data");
            Directory.CreateDirectory(dataDir);

            // Create sample CSV files.
            CreateSampleCsv(Path.Combine(dataDir, "data1.csv"));
            CreateSampleCsv(Path.Combine(dataDir, "data2.csv"));

            // Create the LINQ Reporting template.
            string templatePath = Path.Combine(workDir, "template.docx");
            CreateTemplate(templatePath);

            // Load the template once; it will be cloned for each CSV file.
            Document template = new Document(templatePath);

            // Reporting engine configuration.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Final document that will contain merged results from all CSV files.
            Document finalReport = new Document();

            // Process each CSV file.
            foreach (string csvFile in Directory.GetFiles(dataDir, "*.csv"))
            {
                // Model exposing the file name.
                FileModel model = new FileModel { FileName = Path.GetFileName(csvFile) };

                // CSV data source with header support.
                CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
                {
                    HasHeaders = true,
                    Delimiter = ',',
                    QuoteChar = '"'
                };
                CsvDataSource csvSource = new CsvDataSource(csvFile, loadOptions);

                // Clone the template for the current file.
                Document currentDoc = (Document)template.Clone(true);

                // Build the report using two data sources: the model ("model") and the CSV rows ("rows").
                engine.BuildReport(currentDoc,
                    new object[] { model, csvSource },
                    new[] { "model", "rows" });

                // Append the generated document to the final report.
                finalReport.AppendDocument(currentDoc, ImportFormatMode.KeepSourceFormatting);
            }

            // Save the merged report.
            string outputPath = Path.Combine(workDir, "MergedReport.docx");
            finalReport.Save(outputPath);
        }

        // Generates a simple CSV file with headers Id,Name,Value.
        private static void CreateSampleCsv(string path)
        {
            string[] lines =
            {
                "Id,Name,Value",
                "1,Alpha,10.5",
                "2,Beta,20.75",
                "3,Gamma,30.0"
            };
            File.WriteAllLines(path, lines);
        }

        // Creates a Word template containing LINQ Reporting tags.
        private static void CreateTemplate(string path)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Header showing the source file name.
            builder.Writeln("Report for file: <<[model.FileName]>>");
            builder.Writeln();

            // Table header.
            builder.Writeln("Id\tName\tValue");
            // Begin foreach over CSV rows.
            builder.Writeln("<<foreach [row in rows]>>");
            // Row data.
            builder.Writeln("<<[row.Id]>>\t<<[row.Name]>>\t<<[row.Value]>>");
            // End foreach.
            builder.Writeln("<</foreach>>");
            builder.Writeln();
            builder.Writeln("--------------------------------------------------");
            builder.Writeln();

            doc.Save(path);
        }
    }
}
