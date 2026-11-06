using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
}

public class Model
{
    public List<Person> Persons { get; set; } = new()
    {
        new Person { Name = "Alice", Age = 28 },
        new Person { Name = "Bob", Age = 35 }
    };

    // Returns the JSON representation of the current model instance.
    public string ToJson()
    {
        return JsonConvert.SerializeObject(this);
    }
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new Model();

        // Create a template document with a LINQ Reporting tag that calls the wrapper method.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Serialized JSON:");
        builder.Writeln("<<[model.ToJson()]>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        template.Save("Report.docx");
    }
}
