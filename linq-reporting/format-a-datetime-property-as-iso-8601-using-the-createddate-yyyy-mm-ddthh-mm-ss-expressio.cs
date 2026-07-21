using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model with a DateTime property.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public DateTime CreatedDate { get; set; } = DateTime.Now;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a line that uses a LINQ Reporting expression tag.
            // Use ToString with a custom format to produce ISO‑8601 output.
            builder.Writeln(
                "Created on: <<[model.CreatedDate.ToString(\"yyyy-MM-ddTHH:mm:ss\")]>>");

            // 2. Prepare the data source.
            ReportModel model = new ReportModel
            {
                // Example date; you can set any DateTime you need.
                CreatedDate = new DateTime(2023, 5, 17, 14, 30, 45, DateTimeKind.Utc)
            };

            // 3. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };
            // The third argument is the name used in the template to reference the root object.
            engine.BuildReport(template, model, "model");

            // 4. Save the generated report.
            template.Save("Report.docx");
        }
    }
}
