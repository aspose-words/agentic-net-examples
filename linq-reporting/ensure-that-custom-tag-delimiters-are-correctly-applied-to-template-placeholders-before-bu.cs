using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Replacing;

namespace LinqReportingCustomDelimiters
{
    // Sample data model
    public class Customer
    {
        public string CustomerName { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare working directories
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a template with custom delimiters {{...}}
            string templatePath = Path.Combine(outputDir, "template.docx");
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello {{CustomerName}}!"); // Custom delimiters
            templateDoc.Save(templatePath);

            // 2. Load the template and replace custom delimiters with LINQ Reporting delimiters
            var doc = new Document(templatePath);
            var replaceOptions = new FindReplaceOptions
            {
                MatchCase = false
                // FindWholeWords property does not exist in current API version; default behavior is sufficient.
            };
            doc.Range.Replace("{{", "<<[", replaceOptions);
            doc.Range.Replace("}}", "]>>", replaceOptions);

            // 3. Prepare the data model
            var model = new Customer { CustomerName = "Alice Smith" };

            // 4. Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 5. Save the generated report
            string reportPath = Path.Combine(outputDir, "report.docx");
            doc.Save(reportPath);
        }
    }
}
