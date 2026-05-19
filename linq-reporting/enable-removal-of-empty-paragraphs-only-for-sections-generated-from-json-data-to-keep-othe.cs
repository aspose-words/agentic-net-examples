using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

// Ensure code page support for Aspose.Words
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class Person
{
    public string? Name { get; set; }
    public int Age { get; set; }
}

// Make the class partial to match the test harness definition.
public partial class Program
{
    public static void Main()
    {
        // Paths for temporary files (created in the current working directory)
        string templatePath = "template.docx";
        string jsonPath = "people.json";
        string outputPath = "report.docx";

        // 1. Create sample JSON data (some entries have empty Name to produce empty paragraphs)
        var people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "", Age = 25 },          // Empty name -> empty paragraph after tag removal
            new Person { Name = "Bob", Age = 40 },
            new Person { Name = null, Age = 22 }        // Null name also results in empty paragraph
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(people));

        // 2. Build the template document programmatically
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Static section that must remain untouched
        builder.Writeln("=== Static Section ===");
        builder.Writeln("This paragraph should stay even if it becomes empty after processing.");

        // Insert a section break to separate static content from JSON‑driven content
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // JSON‑driven section
        builder.Writeln("=== JSON Generated Section ===");
        // Begin foreach over the JSON array; the root name will be "persons"
        builder.Writeln("<<foreach [person in persons]>>");
        // The exclamation mark after the tag tells the engine to remove the paragraph if the tag resolves to empty
        builder.Writeln("<<[person.Name]!>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        doc.Save(templatePath);

        // 3. Load the template (demonstrating the load step)
        var loadedTemplate = new Document(templatePath);

        // 4. Prepare the JSON data source
        var jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Configure the reporting engine to remove empty paragraphs
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // 6. Build the report; the root object name must match the tag reference ("persons")
        engine.BuildReport(loadedTemplate, jsonDataSource, "persons");

        // 7. Save the final report
        loadedTemplate.Save(outputPath);

        // Optional cleanup (commented out to keep files for inspection)
        // File.Delete(templatePath);
        // File.Delete(jsonPath);
    }
}
