using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public string Name { get; set; } = "World";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple template document with a LINQ Reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Hello <<[model.Name]>>!");

        // Prepare the root data object.
        Model model = new Model();

        // Create the reporting engine and build the first report.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(template, model, "model");
        Console.WriteLine($"First BuildReport succeeded: {success}");

        // Attempt to call SetRestrictedTypes after the first BuildReport.
        try
        {
            // SetRestrictedTypes is a static method; calling it after BuildReport should throw.
            ReportingEngine.SetRestrictedTypes(new[] { typeof(string) });
            Console.WriteLine("SetRestrictedTypes did NOT throw an exception as expected.");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"Expected exception caught: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected exception type caught: {ex.GetType().Name} - {ex.Message}");
        }

        // Save the generated document to verify output.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        template.Save(outputPath);
        Console.WriteLine($"Output document saved to: {outputPath}");
    }
}
