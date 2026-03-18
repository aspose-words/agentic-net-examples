using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ConditionalReportExample
{
    // Simple data source with a numeric field.
    public class ReportData
    {
        public int Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // 2. Insert a conditional section using Reporting Engine syntax.
            // The section will be displayed only when the numeric field exceeds the threshold (e.g., 100).
            // <<if [Data.Value] > 100>> ... <<endif>>
            builder.Writeln("<<if [Data.Value] > 100>>");
            builder.Writeln("The value is greater than 100.");
            builder.Writeln("<<endif>>");

            // 3. Prepare the data source.
            ReportData data = new ReportData { Value = 150 }; // Change this value to test the condition.

            // 4. Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // Remove empty paragraphs that may remain after the conditional block is omitted.
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // 5. Build the report by merging the template with the data source.
            // The data source name "Data" matches the name used in the template tags.
            engine.BuildReport(template, data, "Data");

            // 6. Save the resulting document.
            template.Save("ConditionalReport.docx");
        }
    }
}
