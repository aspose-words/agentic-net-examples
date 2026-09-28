using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------------------
        // 1. Create a template document using the default tag delimiters << >>
        // -------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Person Report");
        builder.Writeln(); // empty line

        // Start foreach loop.
        builder.Writeln("<<foreach [p in Persons]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Table row with data placeholders.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Age]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------------------
        // 2. Load the template and prepare the reporting engine.
        // -------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // -------------------------------------------------------------
        // 3. Prepare sample data model.
        // -------------------------------------------------------------
        ReportModel model = new()
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // -------------------------------------------------------------
        // 4. Build the report.
        // -------------------------------------------------------------
        bool success = engine.BuildReport(doc, model, "model");

        // -------------------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------------------
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}
