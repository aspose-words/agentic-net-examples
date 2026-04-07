using System;
using System.Collections.Generic;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Sample JSON data
        string json = @"
        [
            { ""Id"": 1, ""Name"": ""Alice"", ""Age"": 30 },
            { ""Id"": 2, ""Name"": ""Bob"", ""Age"": 25 },
            { ""Id"": 3, ""Name"": ""Charlie"", ""Age"": 28 }
        ]";

        // Deserialize JSON into a list of Person objects
        List<Person> persons = JsonConvert.DeserializeObject<List<Person>>(json) ?? new();

        // Wrap the data for the reporting engine
        ReportModel model = new() { Persons = persons };

        // Create a Word document that will serve as the template
        Document template = new();
        DocumentBuilder builder = new(template);

        // Get the properties of Person to generate column headers dynamically
        PropertyInfo[] props = typeof(Person).GetProperties(BindingFlags.Public | BindingFlags.Instance);

        // Begin foreach block – the whole table will be repeated for each person
        builder.Writeln("<<foreach [p in Persons]>>");

        // Start the table
        Table table = builder.StartTable();

        // Header row
        foreach (PropertyInfo prop in props)
        {
            builder.InsertCell();
            builder.Write(prop.Name);
        }
        builder.EndRow();

        // Data row (repeated for each person)
        foreach (PropertyInfo prop in props)
        {
            builder.InsertCell();
            builder.Write($"<<[p.{prop.Name}]>>");
        }
        builder.EndRow();

        // End the table
        builder.EndTable();

        // End foreach block
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection)
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Build the report
        ReportingEngine engine = new();
        bool success = engine.BuildReport(template, model, "model");

        // Save the generated report
        const string outputPath = "Report.docx";
        template.Save(outputPath);

        // Indicate completion
        Console.WriteLine(success
            ? $"Report generated successfully: {Path.GetFullPath(outputPath)}"
            : "Report generation failed.");
    }
}

// Wrapper class for the reporting engine root object
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Sample data model matching the JSON structure
public class Person
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
