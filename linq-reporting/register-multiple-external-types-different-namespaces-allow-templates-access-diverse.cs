using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace MyApp
{
    namespace Models
    {
        public class Customer
        {
            public string Name { get; set; }
            public int Age { get; set; }
        }
    }

    public class DummyData { }

    class Program
    {
        static void Main()
        {
            // Create a minimal document using DocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Report generated.");

            // Initialize the reporting engine
            ReportingEngine engine = new ReportingEngine();

            // Register external types from different namespaces
            engine.KnownTypes.Add(typeof(DateTime));
            engine.KnownTypes.Add(typeof(Math));
            engine.KnownTypes.Add(typeof(Models.Customer));

            // Build the report using a visible data source type
            engine.BuildReport(doc, new DummyData());

            // Save the processed document
            doc.Save("output.docx");
        }
    }
}
