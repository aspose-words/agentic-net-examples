using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputFolder, "template.docx");
        string resultPath = Path.Combine(outputFolder, "result.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple template document with a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back from disk (required by the workflow).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare a data model that matches the tag in the template.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel { Name = "World" };

        // -----------------------------------------------------------------
        // 4. Build the report for the first time – this must succeed.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");
        loadedTemplate.Save(resultPath);

        // -----------------------------------------------------------------
        // 5. Attempt to modify restricted types after the first BuildReport.
        //    According to the documentation this should throw an exception.
        // -----------------------------------------------------------------
        bool exceptionThrown = false;
        try
        {
            // The engine should reject changes to restricted types after a report has been built.
            ReportingEngine.SetRestrictedTypes(typeof(string));
        }
        catch (Exception) // Catch any exception type thrown for this invalid operation.
        {
            exceptionThrown = true;
        }

        // -----------------------------------------------------------------
        // 6. Output the verification result.
        // -----------------------------------------------------------------
        Console.WriteLine($"Exception thrown as expected: {exceptionThrown}");
    }
}
