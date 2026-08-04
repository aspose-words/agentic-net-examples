using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample external type whose properties will be accessed directly from the template.
    public class CustomerInfo
    {
        // Static properties allow direct access without an instance.
        public static string Name { get; } = "John Doe";
        public static int Age { get; } = 42;
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a template document with LINQ Reporting tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Customer Name: <<[CustomerInfo.Name]>>");
            builder.Writeln("Customer Age: <<[CustomerInfo.Age]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Step 2: Load the template for reporting.
            Document loadedTemplate = new Document(templatePath);

            // Step 3: Configure the ReportingEngine and register the external type.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(CustomerInfo));

            // Build the report. No data source is needed because we are accessing static members.
            engine.BuildReport(loadedTemplate, new object(), "");

            // Step 4: Save the generated report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
