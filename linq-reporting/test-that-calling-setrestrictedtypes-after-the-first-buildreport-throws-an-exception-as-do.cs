using System;
using System.IO;
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
        // Prepare file paths in the current working directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // 1. Create a LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!");
        templateDoc.Save(templatePath);

        // 2. Load the template and prepare the data model.
        Document loadedTemplate = new Document(templatePath);
        Model model = new Model();

        // 3. Configure the ReportingEngine and set restricted types BEFORE the first BuildReport.
        ReportingEngine engine = new ReportingEngine();
        ReportingEngine.SetRestrictedTypes(typeof(string)); // static call as required.

        // Build the report for the first time.
        engine.BuildReport(loadedTemplate, model, "model");
        loadedTemplate.Save(outputPath);

        // 4. Attempt to call SetRestrictedTypes AFTER a report has already been built.
        //    According to the documentation this must throw an InvalidOperationException.
        try
        {
            ReportingEngine.SetRestrictedTypes(typeof(int));
            Console.WriteLine("ERROR: No exception was thrown when calling SetRestrictedTypes after BuildReport.");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine("Expected exception caught: " + ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected exception type caught: " + ex.GetType().Name + " - " + ex.Message);
        }
    }
}
