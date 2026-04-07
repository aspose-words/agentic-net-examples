using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // External type to be used inside the template.
    public class MyClass
    {
        public string Name { get; set; } = "World";

        // Static member that can be called from the template.
        public static string GetGreeting()
        {
            return "Hello";
        }
    }

    // Wrapper class that will be passed as the data source.
    public class ReportModel
    {
        public MyClass MyClass { get; set; } = new MyClass();
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Use a static method from MyClass and an instance property.
            builder.Writeln("<<[MyClass.GetGreeting()]>> <<[model.MyClass.Name]>>!");

            // Save the template to a local file.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real‑world scenario).
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            // Enable reflection optimization for faster property access.
            ReportingEngine.UseReflectionOptimization = true;

            ReportingEngine engine = new ReportingEngine();

            // Register the external type so its members can be used in the template.
            engine.KnownTypes.Add(typeof(MyClass));

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                MyClass = new MyClass { Name = "Aspose.Words" }
            };

            // The root object name used in the template is "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
