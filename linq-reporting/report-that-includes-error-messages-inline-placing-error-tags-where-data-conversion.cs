using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorReportDemo
{
    // Simple data class – the property will contain a value that cannot be converted to a number.
    public class SampleData
    {
        public string Amount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document in memory.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // The template contains a field that tries to format the value as a number.
            // When the conversion fails, the ReportingEngine will insert an inline error message
            // surrounded by <<error>> tags because we enable the InlineErrorMessages option.
            builder.Writeln("Amount: <<[data.Amount]:dollarText>>");

            // 2. Prepare a data source with an invalid numeric value.
            SampleData data = new SampleData { Amount = "NotANumber" };

            // 3. Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Enable inline error messages.
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // 4. Build the report. The engine will replace the template tag with either the formatted value
            //    or an inline error message (e.g., <<error: Cannot convert 'NotANumber' to a number>>).
            engine.BuildReport(template, data, "data");

            // 5. Save the resulting document.
            template.Save("InlineErrorReport.docx");
        }
    }
}
