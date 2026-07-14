using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // The tag uses ElementAt(3) to fetch the fourth element (zero‑based index).
        builder.Writeln("Fourth person: <<[model.Persons.ElementAt(3).Name]>> Age: <<[model.Persons.ElementAt(3).Age]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model with at least four persons.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice",   Age = 30 },
                new Person { Name = "Bob",     Age = 25 },
                new Person { Name = "Charlie", Age = 28 },
                new Person { Name = "Diana",   Age = 22 }, // Fourth element (index 3)
                new Person { Name = "Eve",     Age = 35 }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        var outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
