using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Path for the template document.
        const string templatePath = "template.docx";

        // Create a simple template with a LINQ Reporting tag.
        CreateTemplate(templatePath);

        // Load the template document.
        Document template = new Document(templatePath);

        // Initialise the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // First data source.
        var model1 = new Model { Name = "Alice" };
        engine.BuildReport(template, model1, "model");
        template.Save("output1.docx");

        // Reload the template for the second run (the previous document was modified).
        Document template2 = new Document(templatePath);
        var model2 = new Model { Name = "Bob" };
        engine.BuildReport(template2, model2, "model");
        template2.Save("output2.docx");
    }

    // Creates a Word document containing a single reporting tag.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello <<[model.Name]>>!");
        doc.Save(path);
    }

    // Simple data model used by the template.
    public class Model
    {
        public string Name { get; set; } = string.Empty;
    }
}
