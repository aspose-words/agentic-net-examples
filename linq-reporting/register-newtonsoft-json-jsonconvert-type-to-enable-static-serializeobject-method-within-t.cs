using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string FirstName { get; set; } = "John";
    public string LastName { get; set; } = "Doe";
    public int Age { get; set; } = 30;
}

public class Program
{
    public static void Main()
    {
        // Create a simple data model.
        var model = new Person();

        // Create a Word document and insert a LINQ Reporting tag that serializes the model to JSON.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Serialized JSON:");
        // The tag calls the static JsonConvert.SerializeObject method.
        builder.Writeln("<<[JsonConvert.SerializeObject(model)]>>");

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        // Register the JsonConvert type so that static members can be used in the template.
        engine.KnownTypes.Add(typeof(JsonConvert));

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
