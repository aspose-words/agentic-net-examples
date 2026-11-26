using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingAllowMissingMembers
{
    // Sample data model
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
        // Note: No Address property – will be missing in the template.
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider for .NET Core.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the LINQ Reporting template.
            builder.Writeln("Name: <<[model.Name]>>");
            builder.Writeln("Age: <<[model.Age]>>");
            // This field does not exist in the model; AllowMissingMembers will treat it as null.
            builder.Writeln("Address: <<[model.Address]>>");
            // Intentional syntax error: missing closing </if> tag.
            builder.Writeln("<<if [model.Name]>>Hello");

            // Save the template (optional, for inspection).
            string templatePath = "Template.docx";
            doc.Save(templatePath);

            // Load the template back (demonstrates load step).
            Document templateDoc = new Document(templatePath);

            // Prepare the data source.
            Person model = new();

            // Configure the reporting engine with both options.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages;

            // Build the report.
            bool success = engine.BuildReport(templateDoc, model, "model");

            // Save the generated report.
            string outputPath = "Report.docx";
            templateDoc.Save(outputPath);

            // Output the result status.
            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Report saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}
