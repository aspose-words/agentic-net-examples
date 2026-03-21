using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    public class ReportData
    {
        public string CurrentDate { get; set; }
        public string Name { get; set; }
        public int Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a simple template document in memory with Reporting Engine tags.
            Document doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Report generated on <<[Data.CurrentDate]>>");
            builder.Writeln("Hello, <<[Data.Name]>>!");
            builder.Writeln("Value: <<[Data.Value]>>");

            // Enable reflection optimization for faster property access.
            ReportingEngine.UseReflectionOptimization = true;

            // Prepare a data source object with the properties referenced in the template.
            var data = new ReportData
            {
                CurrentDate = DateTime.Now.ToString("yyyy-MM-dd"),
                Name = "World",
                Value = 12345
            };

            // Build the report using the in‑memory data source.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, data, "Data");

            // Save the generated report to the current directory.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Report saved to {outputPath}");
        }
    }
}
