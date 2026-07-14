using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the root object for the report.
    public class SampleModel
    {
        public string Name { get; set; } = "Aspose";
        public int Value { get; set; } = 42;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words for certain encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Tag that uses a static method from System.Math (core assembly).
            // Correct LINQ Reporting syntax for static method call uses dot notation.
            builder.Writeln("Square root of 16 = <<[System.Math.Sqrt(16)]>>");

            // Tag that uses a static method from Newtonsoft.Json (third‑party assembly).
            // Use dot notation for the static method call.
            builder.Writeln("Serialized model = <<[Newtonsoft.Json.JsonConvert.SerializeObject(model)]>>");

            // Tag that accesses a property of the root model object.
            builder.Writeln("Model Name = <<[model.Name]>>");
            builder.Writeln("Model Value = <<[model.Value]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare data.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            SampleModel model = new SampleModel();

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Register core and third‑party types so their static members can be used in the template.
            engine.KnownTypes.Add(typeof(System.Math));
            engine.KnownTypes.Add(typeof(Newtonsoft.Json.JsonConvert));

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            // The root object name used in the template is "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);
        }
    }
}
