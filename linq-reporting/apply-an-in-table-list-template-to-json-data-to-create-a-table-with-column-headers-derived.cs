using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Charlie"", ""Age"": 28 }
        ]";
        File.WriteAllText("data.json", json, Encoding.UTF8);

        // Load JSON into model
        string jsonData = File.ReadAllText("data.json", Encoding.UTF8);
        List<Person> persons = JsonConvert.DeserializeObject<List<Person>>(jsonData) ?? new();
        ReportModel model = new() { Persons = persons };

        // Create template document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin foreach loop for data rows
        builder.Writeln("<<foreach [person in Persons]>>");

        // Build table (header + data row)
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row (repeated for each person)
        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach loop
        builder.Writeln("<</foreach>>");

        // Generate report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the result
        doc.Save("report.docx");

        Console.WriteLine("Report generated: report.docx");
    }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}
