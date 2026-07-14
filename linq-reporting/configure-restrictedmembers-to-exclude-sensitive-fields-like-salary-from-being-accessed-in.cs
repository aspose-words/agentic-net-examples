using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with a sensitive field.
    public class Employee
    {
        public string Name { get; set; } = "John Doe";
        public decimal Salary { get; set; } = 12345.67m;
    }

    // Wrapper class used as the root data source for the report.
    public class ReportModel
    {
        public Employee Employee { get; set; } = new Employee();
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the code page provider is registered (required for some data sources).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags.
            builder.Writeln("Employee Report");
            builder.Writeln("Name: <<[model.Employee.Name]>>");
            builder.Writeln("Salary: <<[model.Employee.Salary]>>"); // Sensitive field.

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure restricted types to block access to sensitive members.
            //    Here we restrict the entire Employee type; any member (including Salary)
            //    will be inaccessible in the template.
            // -----------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(typeof(Employee));

            // -----------------------------------------------------------------
            // 4. Prepare the reporting engine.
            //    AllowMissingMembers prevents exceptions when a restricted member is accessed.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = string.Empty // Hide missing member messages.
            };

            // -----------------------------------------------------------------
            // 5. Build the report.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel(); // Populate with real data as needed.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
