using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data.
        string json = @"
        [
            { ""Name"": ""Alice"", ""Age"": 30, ""Country"": ""USA"" },
            { ""Name"": ""Bob"",   ""Age"": 25, ""Country"": ""Canada"" },
            { ""Name"": ""Charlie"", ""Age"": 35, ""Country"": ""UK"" }
        ]";

        // Deserialize JSON into a list of Person objects.
        List<Person> persons = JsonConvert.DeserializeObject<List<Person>>(json) ?? new();

        // Wrapper model for the report.
        ReportModel model = new()
        {
            Persons = persons
        };

        // Create a Word template programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);

        // -----------------------------------------------------------------
        // Header table (static, appears once)
        // -----------------------------------------------------------------
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln(nameof(Person.Name));
        builder.InsertCell();
        builder.Writeln(nameof(Person.Age));
        builder.InsertCell();
        builder.Writeln(nameof(Person.Country));
        builder.EndRow();

        builder.EndTable();

        // -----------------------------------------------------------------
        // Data rows – generated inside a foreach block
        // -----------------------------------------------------------------
        builder.Writeln("<<foreach [person in Persons]>>");

        // Table that will be repeated for each person.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Country]>>");
        builder.EndRow();

        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Build the report using the template and the model.
        ReportingEngine engine = new();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        const string outputPath = "report.docx";
        template.Save(outputPath);
    }
}

// Data model representing a person.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Country { get; set; } = string.Empty;
}

// Wrapper model for the report.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}
