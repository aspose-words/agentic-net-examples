using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Simple data model used as the root object for the report.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    public static void Main()
    {
        // Register code page provider required by Aspose.Words for some encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Plain data from the root object.
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");

        // Use a static method from System.Guid (built‑in assembly).
        builder.Writeln("Generated GUID: <<[System.Guid.NewGuid()]>>");

        // Use a static method from Newtonsoft.Json (external assembly) to serialize the model.
        builder.Writeln("JSON representation: <<[Newtonsoft.Json.JsonConvert.SerializeObject(model)]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (ensures the template is fully persisted before building).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the root data object.
        // -----------------------------------------------------------------
        Person model = new Person();

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Register external types from multiple assemblies so that the template can access them.
        engine.KnownTypes.Add(typeof(System.Guid));                     // mscorlib / System assembly
        engine.KnownTypes.Add(typeof(Newtonsoft.Json.JsonConvert));    // Newtonsoft.Json assembly

        // -----------------------------------------------------------------
        // 5. Build the report.
        // -----------------------------------------------------------------
        // The root name used in the template tags is "model".
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}
