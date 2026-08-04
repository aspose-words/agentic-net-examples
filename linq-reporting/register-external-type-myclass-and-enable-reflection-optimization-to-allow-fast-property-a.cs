using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample external type to be used in the template.
    public class MyClass
    {
        // Static property accessed via reflection optimization.
        public static string Greeting => "Hello";
    }

    // Simple data model used as the root object for the report.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a line that uses a static member of MyClass and an instance member of Person.
            builder.Writeln("<<[MyClass.Greeting]>> <<[person.Name]>>!");

            // Save the template to disk before building the report.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare data.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            Person person = new Person { Name = "Alice" };

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            // Enable reflection optimization (static property for faster access).
            ReportingEngine.UseReflectionOptimization = true;

            ReportingEngine engine = new ReportingEngine();

            // Register the external type so its static members can be used in the template.
            engine.KnownTypes.Add(typeof(MyClass));

            // Build the report. The root object name must match the tag used in the template.
            engine.BuildReport(doc, person, "person");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);

            // Optional: indicate completion.
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
