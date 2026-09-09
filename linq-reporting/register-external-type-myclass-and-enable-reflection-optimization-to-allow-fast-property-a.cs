using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // External type that will be accessed from the template.
    public class MyClass
    {
        // Static property accessed via the template.
        public static string Greeting => "Hello from MyClass";

        // Static method accessed via the template.
        public static int GetNumber()
        {
            return 42;
        }
    }

    // Root data model for the report.
    public class Model
    {
        // Instance property accessed via the template.
        public string PersonName { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document and a builder to insert content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags that reference the external type and the model.
            builder.Writeln("Greeting: <<[MyClass.Greeting]>>");
            builder.Writeln("Number: <<[MyClass.GetNumber()]>>");
            builder.Writeln("Name: <<[model.PersonName]>>");

            // Save the template to a local file (optional, demonstrates load/save lifecycle).
            const string templatePath = "Template.docx";
            doc.Save(templatePath);

            // Load the template back (simulating a separate load step).
            Document template = new Document(templatePath);

            // Enable reflection optimization for faster property access.
            ReportingEngine.UseReflectionOptimization = true;

            // Create the reporting engine and register the external type.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(MyClass));

            // Prepare the root data object.
            Model model = new Model();

            // Build the report using the template, the model, and the root name "model".
            engine.BuildReport(template, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            template.Save(outputPath);
        }
    }
}
