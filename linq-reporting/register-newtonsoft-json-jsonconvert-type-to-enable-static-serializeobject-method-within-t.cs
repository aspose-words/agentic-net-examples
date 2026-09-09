using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model
        var model = new SampleModel
        {
            Name = "John Doe",
            Age = 30,
            Tags = new List<string> { "Developer", "Blogger", "Speaker" }
        };

        // Create a template document programmatically
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that calls JsonConvert.SerializeObject on the model.
        // The result will be inserted as plain text.
        builder.Writeln("Serialized JSON:");
        builder.Writeln("<<[JsonConvert.SerializeObject(model)]>>");

        // Save the template (optional, just to have a file on disk)
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // Load the template (demonstrates load step)
        var loadedDoc = new Document(templatePath);

        // Configure the reporting engine
        var engine = new ReportingEngine();

        // Register the JsonConvert type to allow static method calls in the template
        engine.KnownTypes.Add(typeof(JsonConvert));

        // Build the report using the model as the root object named "model"
        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report
        const string outputPath = "Report.docx";
        loadedDoc.Save(outputPath);
    }
}

// Sample data model with public properties
public class SampleModel
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public List<string> Tags { get; set; } = new();
}
