using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace MyApp
{
    // Custom type defined in a different namespace to be exposed to templates
    namespace Models
    {
        public class Customer
        {
            public string Name { get; set; }
            public int Age { get; set; }
        }
    }

    class Program
    {
        static void Main()
        {
            // Load a template document (ensure the file exists at the specified path)
            Document doc = new Document("template.docx");

            // Initialize the reporting engine
            ReportingEngine engine = new ReportingEngine();

            // Register multiple external types from different namespaces
            engine.KnownTypes.Add(typeof(System.DateTime));          // System namespace
            engine.KnownTypes.Add(typeof(System.Math));              // System namespace
            engine.KnownTypes.Add(typeof(MyApp.Models.Customer));   // Custom namespace

            // Build the report; using an empty anonymous object as a placeholder data source
            engine.BuildReport(doc, new { });

            // Save the processed document
            doc.Save("output.docx");
        }
    }
}
