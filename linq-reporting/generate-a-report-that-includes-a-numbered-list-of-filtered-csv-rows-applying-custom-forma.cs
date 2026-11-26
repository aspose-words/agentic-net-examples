using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCsvExample
{
    // Model representing a CSV row.
    public class CsvRow
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public string Category { get; set; } = "";
        public double Value { get; set; }
    }

    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public List<CsvRow> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure code page support for CSV parsing if needed.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // 1. Create sample CSV data.
            string csvPath = "data.csv";
            File.WriteAllText(csvPath,
                "Id,Name,Category,Value\n" +
                "1,Alpha,TypeA,45.5\n" +
                "2,Beta,TypeB,78.2\n" +
                "3,Gamma,TypeA,62.0\n" +
                "4,Delta,TypeC,33.1\n" +
                "5,Epsilon,TypeB,91.4\n");

            // 2. Load CSV rows into model, filter rows where Value > 50.
            var model = new ReportModel();
            using (var reader = new StreamReader(csvPath))
            {
                // Skip header.
                string? headerLine = reader.ReadLine();
                int index = 1;
                while (!reader.EndOfStream)
                {
                    string? line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] parts = line.Split(',');
                    if (parts.Length != 4)
                        continue;

                    if (!double.TryParse(parts[3], NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
                        continue;

                    if (value <= 50)
                        continue; // Filter condition.

                    var row = new CsvRow
                    {
                        Index = index++,
                        Name = parts[1],
                        Category = parts[2],
                        Value = value
                    };
                    model.Items.Add(row);
                }
            }

            // 3. Create the LINQ Reporting template programmatically.
            string templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Filtered CSV Items:");
            // Start a numbered list.
            builder.ListFormat.ApplyNumberDefault();

            // Paragraph containing restartNum and foreach tags.
            builder.Writeln("<<restartNum>><<foreach [item in Items]>>");

            // Paragraph that will be repeated for each item.
            builder.Writeln(
                "<<if [item.Index % 2 == 0]>>" +
                "<<backColor [\"LightGray\"]>><<[item.Name]>> - <<[item.Value]>> <</backColor>><</if>>" +
                "<<if [item.Index % 2 != 0]>>" +
                "<<[item.Name]>> - <<[item.Value]>>" +
                "<</if>>");

            // Closing foreach tag.
            builder.Writeln("<</foreach>>");

            // End the numbered list.
            builder.ListFormat.RemoveNumbers();

            // Save the template.
            doc.Save(templatePath);

            // 4. Load the template for report generation.
            var reportDoc = new Document(templatePath);

            // 5. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 6. Save the generated report.
            string outputPath = "report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
