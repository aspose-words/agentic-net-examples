using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that uses Select with a lambda expression.
        // The lambda formats each Person as "Name (Age)" using string concatenation
        // (the $-prefixed interpolated string is not supported in template syntax).
        builder.Writeln("<<foreach [s in Persons.Select(p => p.Name + \" (\" + p.Age + \")\")]>>");
        builder.Writeln("• <<[s]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 45 },
                new Person { Name = "Carol", Age = 27 }
            }
        };

        // Build the report using the model as the root data source named "model".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Wrapper class that matches the root name used in BuildReport.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple data class referenced by the template.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
