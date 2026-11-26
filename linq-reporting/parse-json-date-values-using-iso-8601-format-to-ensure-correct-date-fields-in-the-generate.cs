using System;
using System.Collections.Generic;
using System.IO;
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

        // Sample JSON data with ISO 8601 date strings.
        string json = @"
        [
            { ""Name"": ""Alice"", ""BirthDate"": ""1990-05-15T00:00:00Z"" },
            { ""Name"": ""Bob"",   ""BirthDate"": ""1985-12-03T00:00:00Z"" },
            { ""Name"": ""Carol"", ""BirthDate"": ""2000-07-21T00:00:00Z"" }
        ]";

        // Deserialize JSON into a list of Person objects.
        List<Person> persons = JsonConvert.DeserializeObject<List<Person>>(json) ?? new List<Person>();

        // Prepare the data model for the report.
        ReportModel model = new()
        {
            Persons = persons
        };

        // Create a Word template programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("People Report");
        builder.Writeln("==============");
        builder.Writeln();

        // Begin a foreach loop over the Persons collection.
        builder.Writeln("<<foreach [person in Persons]>>");
        // Output name and formatted birth date.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Birth Date: <<[person.BirthDate.ToString(\"yyyy-MM-dd\")]>>");
        builder.Writeln(); // Blank line between entries.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document doc = new(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}

// Wrapper class that matches the root name used in BuildReport ("model").
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Data model representing a person with a name and a birth date.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public DateTime BirthDate { get; set; }
}
