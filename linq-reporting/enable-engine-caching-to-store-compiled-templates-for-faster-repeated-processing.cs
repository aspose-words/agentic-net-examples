using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used by the template.
    public class Model
    {
        public string Name { get; set; } = "Aspose";
    }

    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that references the model's Name property.
        builder.Writeln("Hello, <<[model.Name]>>!");

        // Save the template to disk (required before building a report).
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Enable engine caching.
        //    The ReportingEngine caches compiled templates automatically when the same
        //    ReportingEngine instance is reused. No additional API is required.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // (Optional) Enable reflection optimization – it does not affect caching
        // but can improve performance for small data sets.
        ReportingEngine.UseReflectionOptimization = true;

        // -----------------------------------------------------------------
        // 4. Build the report the first time.
        // -----------------------------------------------------------------
        Model model = new Model { Name = "World" };
        engine.BuildReport(loadedTemplate, model, "model");
        string firstOutput = Path.Combine(outputDir, "Report1.docx");
        loadedTemplate.Save(firstOutput);

        // -----------------------------------------------------------------
        // 5. Build the report a second time using the same engine instance.
        //    The compiled template is retrieved from the cache, resulting in faster processing.
        // -----------------------------------------------------------------
        // Reload the template to simulate a fresh document instance.
        Document secondTemplate = new Document(templatePath);
        Model secondModel = new Model { Name = "Everyone" };
        engine.BuildReport(secondTemplate, secondModel, "model");
        string secondOutput = Path.Combine(outputDir, "Report2.docx");
        secondTemplate.Save(secondOutput);

        // -----------------------------------------------------------------
        // 6. Inform the user that the reports have been generated.
        // -----------------------------------------------------------------
        Console.WriteLine($"Report 1 saved to: {firstOutput}");
        Console.WriteLine($"Report 2 saved to: {secondOutput}");
    }
}
