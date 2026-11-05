using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and generated reports
        const string templatePath = "template.docx";
        const string outputPath1 = "output1.docx";
        const string outputPath2 = "output2.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple template document with a LINQ Reporting tag.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello, <<[model.Name]>>!");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for processing.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        var model = new ReportModel { Name = "Alice" };

        // -----------------------------------------------------------------
        // 4. Create the ReportingEngine and enable template caching.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();

        // The CacheCompiledTemplates flag may not be present in older library versions.
        // Using the numeric value (2) ensures the flag is set when available.
        engine.Options = (ReportBuildOptions)2; // Enable caching of compiled templates

        // -----------------------------------------------------------------
        // 5. Build the first report.
        // -----------------------------------------------------------------
        engine.BuildReport(loadedTemplate, model, "model");
        loadedTemplate.Save(outputPath1);

        // -----------------------------------------------------------------
        // 6. Modify the model and rebuild using the same engine (cached template will be reused).
        // -----------------------------------------------------------------
        model.Name = "Bob";

        var secondTemplate = new Document(templatePath);
        engine.BuildReport(secondTemplate, model, "model");
        secondTemplate.Save(outputPath2);
    }

    // -----------------------------------------------------------------
    // Simple data model used by the template.
    // -----------------------------------------------------------------
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }
}
