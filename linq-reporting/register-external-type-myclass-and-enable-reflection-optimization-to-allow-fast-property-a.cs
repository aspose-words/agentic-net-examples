using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample class to be used in the template.
    public class MyClass
    {
        public string Name { get; set; } = string.Empty;

        // Static member that can be accessed from the template after registration.
        public static string GetGreeting()
        {
            return "Hello from MyClass!";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags.
            builder.Writeln("Name: <<[data.Name]>>");
            builder.Writeln("Greeting: <<[MyClass.GetGreeting()]>>");

            // Save the template to disk (required before building the report).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            // Enable reflection optimization for faster property access.
            ReportingEngine.UseReflectionOptimization = true;

            ReportingEngine engine = new ReportingEngine();

            // Register the external type so its static members can be used in the template.
            engine.KnownTypes.Add(typeof(MyClass));

            // -----------------------------------------------------------------
            // 4. Prepare the data source.
            // -----------------------------------------------------------------
            MyClass data = new MyClass { Name = "World" };

            // -----------------------------------------------------------------
            // 5. Build the report.
            // -----------------------------------------------------------------
            // The root object name must match the tag prefix used in the template (data).
            engine.BuildReport(reportDoc, data, "data");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
