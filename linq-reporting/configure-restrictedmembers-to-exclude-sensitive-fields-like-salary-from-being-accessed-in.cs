using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model with a sensitive field.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public string Position { get; set; } = string.Empty;
        public decimal Salary { get; set; }
    }

    // Wrapper model used as the root object for the report.
    public class ReportModel
    {
        public Employee Emp { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Prepare folders.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Employee Report");
            builder.Writeln("Name: <<[model.Emp.Name]>>");
            builder.Writeln("Position: <<[model.Emp.Position]>>");
            // Salary is a sensitive field that we want to hide.
            builder.Writeln("Salary: <<[model.Emp.Salary]>>");

            string templatePath = Path.Combine(outputDir, "Template.docx");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (required before building the report).
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure restricted types to block access to Employee members.
            //    This will make all Employee members inaccessible in the template.
            //    For demonstration we will allow missing members so the engine
            //    does not throw an exception.
            // -----------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(typeof(Employee));

            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = string.Empty // hide missing member messages.
            };

            // -----------------------------------------------------------------
            // 4. Prepare the data source.
            // -----------------------------------------------------------------
            Employee employee = new Employee
            {
                Name = "John Doe",
                Position = "Software Engineer",
                Salary = 95000m
            };

            ReportModel model = new ReportModel { Emp = employee };

            // -----------------------------------------------------------------
            // 5. Build the report. The root object name is "model".
            // -----------------------------------------------------------------
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(outputPath);

            // Optional: output the plain text to the console to verify the result.
            Console.WriteLine("Report generated. Document text:");
            Console.WriteLine(doc.GetText());
        }
    }
}
