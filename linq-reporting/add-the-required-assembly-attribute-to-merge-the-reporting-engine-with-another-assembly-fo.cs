using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

[assembly: Aspose.Words.Reporting.ReportingEngineAssemblyMerge]

namespace LinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Sample data model
            var model = new ReportModel { Name = "World" };

            // Create template document with a LINQ Reporting tag
            const string templatePath = "template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello <<[model.Name]>>!");
            templateDoc.Save(templatePath);

            // Load the template for reporting
            var doc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputPath = "output.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Report saved to: {Path.GetFullPath(outputPath)}");
        }
    }

    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }
}

// Dummy attribute definition to satisfy compilation if the real attribute is unavailable
namespace Aspose.Words.Reporting
{
    [AttributeUsage(AttributeTargets.Assembly)]
    public sealed class ReportingEngineAssemblyMergeAttribute : Attribute
    {
        public ReportingEngineAssemblyMergeAttribute() { }
    }
}
