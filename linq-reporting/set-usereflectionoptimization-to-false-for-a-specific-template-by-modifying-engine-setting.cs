using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Sample data model with a non‑nullable property initialized to avoid warnings.
    public string Name { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a simple template document containing a LINQ Reporting tag.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!"); // Tag will be replaced by the model value.

        // Save the template to a local file.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back from disk.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source (root object) for the report.
        // -----------------------------------------------------------------
        var model = new Model { Name = "World" };

        // -----------------------------------------------------------------
        // 4. Build the report inside a using block while disabling reflection optimization.
        // -----------------------------------------------------------------
        using (var dummyStream = new MemoryStream()) // Using block as required; the stream itself is not used.
        {
            // Disable the reflection optimization for this report generation.
            ReportingEngine.UseReflectionOptimization = false;

            // Create the reporting engine and build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");
        }

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}
