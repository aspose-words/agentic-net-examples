using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model.
        var model = new SampleModel
        {
            Id = 1,
            Name = "John Doe",
            Tags = new List<string> { "example", "json", "serialization" }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Serialized JSON:");
        // The tag calls the static JsonConvert.SerializeObject method.
        builder.Writeln("<<[JsonConvert.SerializeObject(model)]>>");
        doc.Save(templatePath);

        // Load the template document.
        var loadedDoc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        // Register the JsonConvert type to allow static method calls in the template.
        engine.KnownTypes.Add(typeof(JsonConvert));

        // Build the report using the model as the root object named "model".
        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report.
        loadedDoc.Save("Report.docx");
    }
}

// Sample data model used in the report.
public class SampleModel
{
    public int Id { get; set; } = 0;
    public string Name { get; set; } = string.Empty;
    public List<string> Tags { get; set; } = new();
}
