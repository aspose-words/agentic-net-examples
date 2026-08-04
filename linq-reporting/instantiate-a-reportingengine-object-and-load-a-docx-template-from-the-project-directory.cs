using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Public property accessed by the template tag <<[model.Name]>>
    public string Name { get; set; } = "World";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Path to the template file in the project directory
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");

        // Create the template if it does not already exist
        if (!File.Exists(templatePath))
        {
            // Create a blank document and add a simple LINQ Reporting tag
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello, <<[model.Name]>>!");
            // Save the template to disk
            templateDoc.Save(templatePath);
        }

        // Load the template document
        Document doc = new Document(templatePath);

        // Prepare the data source
        Model model = new Model { Name = "Aspose" };

        // Instantiate the reporting engine
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model as the root object named "model"
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);
    }
}
