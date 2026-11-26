using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some data sources)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create a new document and build the template
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Add a heading
        builder.Writeln("Person Report");
        builder.Writeln();

        // Insert an empty paragraph that should be removed
        builder.Writeln();

        // LINQ Reporting foreach block
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection)
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // Load the template back (demonstrates load step)
        var loadedDoc = new Document(templatePath);

        // Configure the reporting engine with RemoveEmptyParagraphs option
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report
        engine.BuildReport(loadedDoc, model, "model");

        // Save the final report
        const string outputPath = "Report.docx";
        loadedDoc.Save(outputPath);
    }
}
