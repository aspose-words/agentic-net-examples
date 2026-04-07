using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a simple template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // Insert a LINQ Reporting tag that will be replaced with the model's Name property.
        builder.Writeln("<<[model.Name]>>");
        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back for reporting.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Define a restricted type list BEFORE the first BuildReport call.
        // -----------------------------------------------------------------
        // For demonstration we restrict the System.String type.
        ReportingEngine.SetRestrictedTypes(typeof(string));

        // -----------------------------------------------------------------
        // 4. Build the report using a simple data model.
        // -----------------------------------------------------------------
        var model = new ReportModel { Name = "Aspose.Words" };
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);

        // -----------------------------------------------------------------
        // 5. Verify that the restricted type list is now immutable.
        //    Attempting to modify it should throw an InvalidOperationException.
        // -----------------------------------------------------------------
        try
        {
            // This call must fail because the restricted types have already been locked.
            ReportingEngine.SetRestrictedTypes(typeof(int));
            Console.WriteLine("ERROR: Restricted types were modified after BuildReport – test failed.");
        }
        catch (InvalidOperationException ex)
        {
            // Expected path – the list is immutable.
            Console.WriteLine("Restricted types are immutable after BuildReport. Caught exception:");
            Console.WriteLine(ex.Message);
        }

        // Indicate completion.
        Console.WriteLine("Report generated at: " + reportPath);
    }
}

// Simple data model used by the template.
public class ReportModel
{
    // Initialise to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}
