using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    // Wrapper class used as the root data source for the report.
    public class ReportModel
    {
        public Person Person { get; set; } = new Person();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document and add a template tag that calls JsonConvert.SerializeObject.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Serialized JSON: <<[JsonConvert.SerializeObject(Person)]>>");

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // Register the JsonConvert type so that static methods can be used in the template.
            engine.KnownTypes.Add(typeof(JsonConvert));

            // Build the report. No data source name is required because we reference members directly.
            engine.BuildReport(doc, model);

            // Save the generated document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
            doc.Save(outputPath);
        }
    }
}
