using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace LinqReportingJsonFilter
{
    // Model representing a single JSON record.
    public class Record
    {
        public int Id { get; set; } = 0;
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; } = 0;
        public string Country { get; set; } = string.Empty;
        public bool IsActive { get; set; } = false;
        public double Score { get; set; } = 0.0;
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
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for temporary files.
            string dataPath = "data.json";
            string templatePath = "template.docx";
            string outputPath = "report.docx";

            // 1. Create sample JSON data.
            var sampleData = new List<Record>
            {
                new Record { Id = 1, Name = "Alice", Age = 28, Country = "USA", IsActive = true, Score = 85.5 },
                new Record { Id = 2, Name = "Bob", Age = 35, Country = "USA", IsActive = false, Score = 78.0 },
                new Record { Id = 3, Name = "Charlie", Age = 42, Country = "Canada", IsActive = true, Score = 92.3 },
                new Record { Id = 4, Name = "Diana", Age = 31, Country = "USA", IsActive = true, Score = 88.1 },
                new Record { Id = 5, Name = "Ethan", Age = 27, Country = "UK", IsActive = true, Score = 73.4 }
            };
            File.WriteAllText(dataPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

            // 2. Load JSON data into objects.
            var allRecords = JsonConvert.DeserializeObject<List<Record>>(File.ReadAllText(dataPath)) ?? new List<Record>();

            // 3. Apply a complex predicate using Where.
            var filteredRecords = allRecords.Where(r =>
                r.Age > 30 &&
                r.Country == "USA" &&
                r.IsActive &&
                r.Score >= 80.0);

            // 4. Create the LINQ Reporting template programmatically.
            var builder = new DocumentBuilder();
            builder.Writeln("Filtered Records:");
            builder.Writeln("<<foreach [rec in Records]>>");
            builder.Writeln("- <<[rec.Id]>>: <<[rec.Name]>> (Age: <<[rec.Age]>>, Country: <<[rec.Country]>>, Active: <<[rec.IsActive]>>, Score: <<[rec.Score]>>)");
            builder.Writeln("<</foreach>>");
            builder.Writeln("End of report.");
            builder.Document.Save(templatePath);

            // 5. Load the template document.
            var doc = new Document(templatePath);

            // 6. Prepare the model with filtered data.
            var model = new ReportModel { Records = filteredRecords.ToList() };

            // 7. Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 8. Save the generated report.
            doc.Save(outputPath);
        }
    }
}
