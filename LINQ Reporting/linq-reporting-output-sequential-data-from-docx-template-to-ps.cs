using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Prepare the data source – a list of simple objects.
        var people = new List<Person>
        {
            new Person { Name = "Alice",   Age = 30 },
            new Person { Name = "Bob",     Age = 25 },
            new Person { Name = "Charlie", Age = 35 }
        };

        // Create a DOCX template that uses LINQ Reporting syntax.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        // The foreach tag iterates over the "people" collection.
        // Inside the loop we output each person's Name and Age.
        builder.Writeln("<<foreach [people]>>");
        builder.Writeln("<<[Name]>> - <<[Age]>>");
        builder.Writeln("<</foreach>>");

        // Populate the template with the data source.
        ReportingEngine engine = new ReportingEngine();
        // The anonymous object provides a property named "people" that the template can reference.
        engine.BuildReport(template, new { people });

        // Save the resulting document as PostScript.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps
        };
        template.Save("Report.ps", psOptions);
    }

    // Simple data class used in the example.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
