using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model for each record.
    public class ReportItem
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    // Custom data source that streams items lazily.
    public class LargeDataSource : IEnumerable<ReportItem>
    {
        private readonly int _count;

        public LargeDataSource(int count = 10000) => _count = count;

        public IEnumerator<ReportItem> GetEnumerator()
        {
            for (int i = 1; i <= _count; i++)
                yield return new ReportItem { Id = i, Name = $"Item {i}" };
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a blank Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the template with LINQ Reporting tags.
            builder.Writeln("Report of items:");
            builder.Writeln("<<foreach [item in data]>>");
            builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Prepare the custom data source.
            LargeDataSource dataSource = new LargeDataSource();

            // Configure and run the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default behavior
            bool success = engine.BuildReport(doc, dataSource, "data");

            // Save the generated report.
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine(success
                ? $"Report generated successfully: {Path.GetFullPath(outputPath)}"
                : "Report generation failed.");
        }
    }
}
