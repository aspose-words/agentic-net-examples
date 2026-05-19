using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "Alice";
    public int Age { get; set; } = 30;
}

public class ReportModel
{
    public Person Person { get; set; } = new Person();
}

public class Program
{
    public static void Main()
    {
        // Create a simple template document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Serialized JSON:");
        // Use a LINQ Reporting tag that calls the static JsonConvert.SerializeObject method.
        builder.Writeln("<<[JsonConvert.SerializeObject(model.Person)]>>");

        // Prepare the data model.
        ReportModel model = new ReportModel();

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register the JsonConvert type so its static methods can be used in the template.
        engine.KnownTypes.Add(typeof(JsonConvert));

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportOutput.docx");
        doc.Save(outputPath);
    }
}
