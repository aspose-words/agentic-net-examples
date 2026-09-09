using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags into the template.
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 45 }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Serialize the generated report to a memory stream (e.g., for emailing).
        using (MemoryStream reportStream = new MemoryStream())
        {
            doc.Save(reportStream, SaveFormat.Docx);
            reportStream.Position = 0; // Reset position for downstream consumers.

            // Example usage: display the size of the generated report.
            Console.WriteLine($"Report generated. Stream length: {reportStream.Length} bytes.");
        }
    }
}

// Wrapper class that matches the root object name used in BuildReport.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple data model referenced by the template tags.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
