using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Model class that holds the HTML snippet to be inserted.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string HtmlSnippet { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank Word document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // 2. Insert the LINQ Reporting tag that uses the -html switch.
            //    The expression evaluates to the HtmlSnippet property of the model.
            builder.Writeln("<<[model.HtmlSnippet] -html>>");

            // 3. Prepare the data model with a formatted HTML fragment.
            ReportModel model = new ReportModel
            {
                HtmlSnippet = @"
                    <h2 style='color:steelblue;'>Report Title</h2>
                    <p>This paragraph is <b>bold</b> and this one is <i>italic</i>.</p>
                    <ul>
                        <li>First item</li>
                        <li>Second item</li>
                    </ul>"
            };

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(template, model, "model");

            // 5. Save the generated document.
            const string outputPath = "Report.docx";
            template.Save(outputPath);
            Console.WriteLine($"Report generated and saved to '{outputPath}'.");
        }
    }
}
