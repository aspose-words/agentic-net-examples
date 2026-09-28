using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and final report.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Simple Person Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (ensures template is fully loaded before building).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Persons = new()
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 25 }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
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
