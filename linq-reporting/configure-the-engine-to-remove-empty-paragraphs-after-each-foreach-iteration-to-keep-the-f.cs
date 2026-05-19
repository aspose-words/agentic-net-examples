using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document with a foreach loop.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a heading.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin foreach loop over the collection "Persons".
        builder.Writeln("<<foreach [p in Persons]>>");

        // Paragraph that may become empty if Name is null or empty.
        builder.Writeln("Name: <<[p.Name]>>");

        // Paragraph that may become empty if Age is null (int is non‑nullable, so we use a sentinel value).
        builder.Writeln("Age: <<[p.Age]>>");

        // End of foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare data.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data: one person with full data, one with empty name.
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "", Age = 0 }   // This will produce empty paragraphs.
            }
        };

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine to remove empty paragraphs.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report. The root object name must match the tag prefix used in the template.
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the final document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public with public properties, non‑nullable init).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
