using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace BatchCsvLinqReporting
{
    // Data model for a single CSV record.
    public class Record
    {
        public int Id { get; set; } = 0;
        public string Name { get; set; } = "";
        public double Value { get; set; } = 0.0;
    }

    // Wrapper model passed to the reporting engine.
    public class ReportModel
    {
        public List<Record> Records { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV reading (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare folders.
            string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "data");
            string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(dataFolder);
            Directory.CreateDirectory(outputFolder);

            // Create sample CSV files.
            CreateSampleCsv(Path.Combine(dataFolder, "file1.csv"));
            CreateSampleCsv(Path.Combine(dataFolder, "file2.csv"));

            // Load and merge CSV data.
            var allRecords = new List<Record>();
            foreach (string csvPath in Directory.GetFiles(dataFolder, "*.csv"))
            {
                LoadCsvIntoList(csvPath, allRecords);
            }

            // Prepare the model.
            var model = new ReportModel { Records = allRecords };

            // Create the template document programmatically.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
            CreateTemplate(templatePath);

            // Load the template.
            var doc = new Document(templatePath);

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the final report.
            string reportPath = Path.Combine(outputFolder, "MergedReport.docx");
            doc.Save(reportPath);
        }

        // Writes a simple CSV file with a header and three rows.
        private static void CreateSampleCsv(string path)
        {
            var lines = new List<string>
            {
                "Id,Name,Value",
                "1,Alpha,10.5",
                "2,Beta,20.75",
                "3,Gamma,30.0"
            };
            File.WriteAllLines(path, lines, Encoding.UTF8);
        }

        // Parses a CSV file and adds its records to the provided list.
        private static void LoadCsvIntoList(string csvPath, List<Record> target)
        {
            var lines = File.ReadAllLines(csvPath, Encoding.UTF8);
            if (lines.Length < 2) return; // No data.

            // Assume first line is header.
            for (int i = 1; i < lines.Length; i++)
            {
                var parts = lines[i].Split(',');
                if (parts.Length != 3) continue;

                if (int.TryParse(parts[0], out int id) &&
                    double.TryParse(parts[2], out double value))
                {
                    var record = new Record
                    {
                        Id = id,
                        Name = parts[1],
                        Value = value
                    };
                    target.Add(record);
                }
            }
        }

        // Generates a Word template containing a foreach loop over Records.
        private static void CreateTemplate(string path)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Merged CSV Report");
            builder.Writeln("<<foreach [rec in Records]>>");
            builder.Writeln("Id: <<[rec.Id]>>, Name: <<[rec.Name]>>, Value: <<[rec.Value]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(path);
        }
    }
}
