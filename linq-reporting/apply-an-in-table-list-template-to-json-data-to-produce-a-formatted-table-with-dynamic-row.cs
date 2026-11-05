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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        var jsonData = new ReportData
        {
            Persons = new List<Person>
            {
                new Person { Index = 1, Name = "Alice" },
                new Person { Index = 2, Name = "Bob" },
                new Person { Index = 3, Name = "Charlie" }
            }
        };
        string jsonString = JsonConvert.SerializeObject(jsonData, Formatting.Indented);
        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, jsonString, Encoding.UTF8);

        // Create the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach loop over Persons collection.
        builder.Writeln("<<foreach [person in Persons]>>");

        // Create a table that will be repeated for each person.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[person.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template back for report generation.
        var doc = new Document(templatePath);

        // Deserialize JSON data into a strongly typed model.
        var model = JsonConvert.DeserializeObject<ReportData>(File.ReadAllText(jsonPath, Encoding.UTF8))!;

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Root data model.
public class ReportData
{
    public List<Person> Persons { get; set; } = new();
}

// Person item model.
public class Person
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
