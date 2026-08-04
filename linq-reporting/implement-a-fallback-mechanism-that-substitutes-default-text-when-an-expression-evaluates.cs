using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Nullable property to allow a null value for demonstration.
    public string? Name { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputFolder);

        // Paths for the template and the final report.
        string templatePath = Path.Combine(outputFolder, "template.docx");
        string resultPath = Path.Combine(outputFolder, "result.docx");

        // -------------------- Create template document --------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a LINQ Reporting tag that will be replaced by the model's Name.
        // If Name is null, the fallback text set in MissingMemberMessage will be used.
        builder.Writeln("Customer name: <<[model.Name]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------- Load template and build report --------------------
        Document reportDoc = new Document(templatePath);

        // Configure the reporting engine to treat missing members (null values) as empty literals.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "N/A"; // Fallback text for null expressions.

        // Create a data model where Name is intentionally null.
        Model model = new Model { Name = null };

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(resultPath);

        // Inform the user where the report was saved.
        Console.WriteLine($"Report generated at: {resultPath}");
    }
}
