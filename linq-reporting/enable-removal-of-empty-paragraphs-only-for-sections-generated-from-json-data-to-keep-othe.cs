using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some JSON scenarios).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // File paths.
        string templatePath = "template.docx";
        string jsonPath = "persons.json";
        string outputPath = "report.docx";

        // ---------- Create sample JSON data ----------
        var persons = new List<Person>
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "", Age = 25 },          // Empty name – should cause paragraph removal.
            new Person { Name = "Bob", Age = 40 }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(persons));

        // ---------- Build the template document ----------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Static section (will not be affected by empty‑paragraph removal).
        builder.Writeln("=== Static Section ===");
        builder.Writeln("<<[model.StaticText]>>"); // Normal tag, no exclamation.

        // Start a new section that will be populated from JSON data.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("=== JSON‑Generated Section ===");
        builder.Writeln("<<foreach [person in model.Persons]>>");

        // Tag with exclamation mark – empty paragraphs resulting from this tag will be removed.
        // The exclamation mark must be placed **after** the closing tag delimiters.
        builder.Writeln("Name: <<[person.Name]>>!"); // If Name is empty, the whole paragraph is removed.
        builder.Writeln("Age: <<[person.Age]>>");

        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // ---------- Load the template ----------
        var template = new Document(templatePath);

        // ---------- Prepare the data model ----------
        var model = new ReportModel
        {
            StaticText = "This is static content.",
            Persons = JsonConvert.DeserializeObject<List<Person>>(File.ReadAllText(jsonPath))
        };

        // ---------- Configure and run the reporting engine ----------
        var engine = new ReportingEngine
        {
            // Remove empty paragraphs only where tags contain an exclamation mark.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(template, model, "model");

        // ---------- Save the final report ----------
        template.Save(outputPath);
    }
}

// Data model for the report.
public class ReportModel
{
    public string StaticText { get; set; } = string.Empty;
    public List<Person> Persons { get; set; } = new();
}

// Simple POCO representing a person.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
