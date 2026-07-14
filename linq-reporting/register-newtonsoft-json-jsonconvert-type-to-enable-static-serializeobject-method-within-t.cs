using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Sample data model
        var person = new Person { Name = "John Doe", Age = 30 };

        // Create a template document with a tag that serializes the root object to JSON
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Serialized JSON:");
        builder.Writeln("<<[JsonConvert.SerializeObject(model)]>>");

        // Configure the reporting engine
        var engine = new ReportingEngine();
        // Register JsonConvert to allow static method calls in the template
        engine.KnownTypes.Add(typeof(JsonConvert));

        // Build the report using the model as the root object named "model"
        engine.BuildReport(doc, person, "model");

        // Save the generated report
        doc.Save("Report.docx");
    }
}
