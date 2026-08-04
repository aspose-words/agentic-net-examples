using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample data model.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    // Utility class whose static members will be accessed from the template.
    public static class MyUtils
    {
        public static string ToUpper(string value) => value?.ToUpperInvariant() ?? string.Empty;
        public static string FormatAge(int age) => $"Age: {age}";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string resultPath = Path.Combine(outputDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write simple placeholders using correct LINQ Reporting syntax.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Uppercase Name: <<[MyUtils.ToUpper(person.Name)]>>");
            builder.Writeln("Age Info: <<[MyUtils.FormatAge(person.Age)]>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for reporting.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Register the external type so its static members can be used safely.
            engine.KnownTypes.Add(typeof(MyUtils));

            // -----------------------------------------------------------------
            // 4. Build the report using a root object named "person".
            // -----------------------------------------------------------------
            Person person = new Person { Name = "Alice Smith", Age = 42 };
            engine.BuildReport(doc, person, "person");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(resultPath);
        }
    }
}
