using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Path to the external Word document to be inserted.
    public string ExternalDocPath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Create an output folder for all generated files.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Step 1: Create the external Word document that will be inserted.
        // -----------------------------------------------------------------
        string externalDocPath = Path.Combine(outputFolder, "External.docx");
        Document externalDoc = new Document();
        DocumentBuilder externalBuilder = new DocumentBuilder(externalDoc);
        externalBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save(externalDocPath);

        // ---------------------------------------------------------------
        // Step 2: Create the template document with a placeholder tag.
        // ---------------------------------------------------------------
        string templatePath = Path.Combine(outputFolder, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Report Start ===");
        // Placeholder that will be replaced by the external document.
        builder.Writeln("<<doc [model.ExternalDocPath]>>");
        builder.Writeln("=== Report End ===");

        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // Step 3: Load the template and build the report.
        // ---------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Prepare the data model with the path to the external document.
        ReportModel model = new ReportModel { ExternalDocPath = externalDocPath };

        // Use the LINQ Reporting engine to process the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // ---------------------------------------------------------------
        // Step 4: Save the final report.
        // ---------------------------------------------------------------
        string resultPath = Path.Combine(outputFolder, "Result.docx");
        loadedTemplate.Save(resultPath);

        // The example finishes without waiting for user input.
    }
}
