using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new Person();

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags.
        builder.Writeln("Static Math.PI value: <<[Math.PI]>>");
        builder.Writeln("Person name: <<[model.Name]>>");
        builder.Writeln("Person age: <<[model.Age]>>");
        builder.Writeln("Serialized JSON (HTML escaped): <<[JsonConvert.SerializeObject(model)] -html>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();

        // Register core and third‑party types for static member access.
        engine.KnownTypes.Add(typeof(Math));               // Core .NET type.
        engine.KnownTypes.Add(typeof(JsonConvert));        // Third‑party type from Newtonsoft.Json.

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
