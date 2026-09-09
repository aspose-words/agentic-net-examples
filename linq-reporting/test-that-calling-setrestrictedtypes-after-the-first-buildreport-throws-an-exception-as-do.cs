using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Create a simple template with a LINQ Reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Hello <<[model.Name]>>!");

        // Save the template to disk.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        template.Save(templatePath);

        // Load the template back.
        Document doc = new Document(templatePath);

        // Prepare the data source.
        Model model = new Model { Name = "World" };

        // Build the first report – this must succeed.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);

        // Attempt to set restricted types after the first BuildReport.
        try
        {
            // This call should throw an exception because restricted types cannot be changed after building a report.
            ReportingEngine.SetRestrictedTypes(typeof(string));
            Console.WriteLine("SetRestrictedTypes did not throw an exception (unexpected).");
        }
        catch (ArgumentException ex)
        {
            // Documented exception type.
            Console.WriteLine($"Caught expected ArgumentException: {ex.Message}");
        }
        catch (InvalidOperationException ex)
        {
            // Actual exception type thrown by the current library version.
            Console.WriteLine($"Caught expected InvalidOperationException: {ex.Message}");
        }
    }
}
