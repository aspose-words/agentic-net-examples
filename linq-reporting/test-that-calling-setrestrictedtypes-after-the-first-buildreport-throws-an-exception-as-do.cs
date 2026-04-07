using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used by the template.
    public class Model
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "Aspose";
    }

    public static void Main()
    {
        // Prepare a temporary folder for the files created by the example.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path of the template document.
        string templatePath = Path.Combine(outputDir, "template.docx");
        // Path of the generated report (optional, just to demonstrate saving).
        string reportPath = Path.Combine(outputDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create a minimal Word template containing a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        // The tag references the root data source named "model".
        builder.Writeln("Hello <<[model.Name]>>!");
        // Save the template to disk.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the first report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        Model model = new Model();

        ReportingEngine engine = new ReportingEngine();
        // BuildReport populates the document using the model object.
        engine.BuildReport(doc, model, "model");
        // Save the generated report (optional, just to have an output file).
        doc.Save(reportPath);

        // -----------------------------------------------------------------
        // 3. Attempt to modify restricted types after the first BuildReport.
        //    According to the documentation this must throw an ArgumentException.
        // -----------------------------------------------------------------
        try
        {
            // This call is illegal after BuildReport has been executed.
            ReportingEngine.SetRestrictedTypes(typeof(string));
            // If no exception is thrown, indicate unexpected behavior.
            Console.WriteLine("Error: No exception was thrown when calling SetRestrictedTypes after BuildReport.");
        }
        catch (ArgumentException ex)
        {
            // Expected path – the engine throws ArgumentException.
            Console.WriteLine("Caught expected ArgumentException: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Any other exception type is unexpected.
            Console.WriteLine("Caught unexpected exception type: " + ex.GetType().Name + " - " + ex.Message);
        }
    }
}
