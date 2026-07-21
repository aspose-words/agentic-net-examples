using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Original data model (can still be used elsewhere).
    public class CustomerModel
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    // Read‑only wrapper exposing only getters – this is what the template can access.
    public class CustomerModelReadOnly
    {
        public string Name { get; }
        public int Age { get; }

        public CustomerModelReadOnly(CustomerModel source)
        {
            Name = source.Name;
            Age = source.Age;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a template document programmatically.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Customer Report");
            builder.Writeln("Name: <<[model.Name]>>");
            builder.Writeln("Age: <<[model.Age]>>");
            // Save the template to disk.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            templateDoc.Save(templatePath);

            // 2. Load the template document (simulating a separate load step).
            Document doc = new Document(templatePath);

            // 3. Prepare the data source.
            CustomerModel model = new CustomerModel();

            // Wrap the model in a read‑only view so that only getters are reachable from the template.
            CustomerModelReadOnly readOnlyModel = new CustomerModelReadOnly(model);

            // 4. Configure the ReportingEngine (no RestrictedMembers property exists).
            ReportingEngine engine = new ReportingEngine();

            // 5. Build the report using the root object name "model".
            engine.BuildReport(doc, readOnlyModel, "model");

            // 6. Save the generated report.
            string resultPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(resultPath);

            Console.WriteLine($"Report generated: {resultPath}");
        }
    }
}
