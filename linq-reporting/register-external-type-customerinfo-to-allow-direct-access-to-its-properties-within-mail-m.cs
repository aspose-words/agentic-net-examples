using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample external type whose static members will be accessed from the template.
    public static class CustomerInfo
    {
        public static string Name => "John Doe";
        public static int Age => 30;
        public static string Email => "john.doe@example.com";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Customer Report");
            builder.Writeln("Name: <<[CustomerInfo.Name]>>");
            builder.Writeln("Age: <<[CustomerInfo.Age]>>");
            builder.Writeln("Email: <<[CustomerInfo.Email]>>");

            // Save the template to a local file.
            const string templatePath = "CustomerTemplate.docx";
            template.Save(templatePath);

            // Load the template for reporting.
            Document report = new Document(templatePath);

            // Configure the ReportingEngine and register the external type.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(CustomerInfo));

            // Build the report. No data source is required because we only use static members.
            engine.BuildReport(report, new object(), "");

            // Save the generated report.
            const string outputPath = "CustomerReport.docx";
            report.Save(outputPath);
        }
    }
}
