using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample JSON data
        string json = @"{ ""Persons"": [ { ""Name"": ""Alice"", ""Age"": 30 }, { ""Name"": ""Bob"", ""Age"": 25 } ] }";

        // Deserialize JSON into the strongly‑typed model
        Model model = JsonConvert.DeserializeObject<Model>(json) ?? new Model();

        // Build the template document programmatically
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Persons Report:");
        builder.Writeln("<<foreach [person in model.Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report to a memory stream
        using MemoryStream ms = new MemoryStream();
        doc.Save(ms, SaveFormat.Docx);
        Console.WriteLine($"Report generated, size: {ms.Length} bytes");
    }
}
