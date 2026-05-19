using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // NOTE: The original example attempted to use an assembly attribute
    // Aspose.Words.Reporting.CustomAuthenticationModule which is not present
    // in the referenced Aspose.Words packages. The attribute has been omitted
    // to allow the code to compile and run. If the attribute becomes available
    // in a newer version, it can be re‑added as:
    // [assembly: Aspose.Words.Reporting.CustomAuthenticationModuleAttribute(typeof(MyAuthenticationModule))]
    public class MyAuthenticationModule
    {
        // This method will be called by the reporting engine when it needs to authenticate a resource.
        public string GetAuthentication(string uri)
        {
            // Return a dummy token or credentials.
            return "Bearer dummy-token";
        }
    }

    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        public string Title { get; set; } = "Sample Report";
        public string Content { get; set; } = "Hello from LINQ Reporting.";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document programmatically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<<[model.Content]>>");

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            ReportModel model = new ReportModel();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("ReportOutput.docx");
        }
    }
}
