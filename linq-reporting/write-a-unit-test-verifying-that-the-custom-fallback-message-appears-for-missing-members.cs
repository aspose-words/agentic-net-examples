using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple template with a missing member reference
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.MissingMember]>>!");
        templateDoc.Save(templatePath);

        // Load the template (simulating a real scenario)
        Document doc = new Document(templatePath);

        // Prepare model without the referenced member
        var model = new Model();

        // Configure the reporting engine to use inline error messages
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report
        bool success = engine.BuildReport(doc, model, "model");

        // Retrieve the generated text
        string resultText = doc.GetText();

        // Determine if the fallback message appears (Aspose.Words inserts "??" for missing members)
        bool containsFallback = resultText.Contains("??") || resultText.Contains("Error") || resultText.Contains("missing");

        // Simple unit‑test style output
        if (!success && containsFallback)
        {
            Console.WriteLine("Test Passed: Fallback message detected.");
        }
        else
        {
            Console.WriteLine("Test Failed: Fallback message not detected.");
        }

        // Save the generated document for inspection (optional)
        doc.Save("output.docx");
    }

    // Model class without the MissingMember property
    public class Model
    {
        public string ExistingMember { get; set; } = "ExistingValue";
    }
}
