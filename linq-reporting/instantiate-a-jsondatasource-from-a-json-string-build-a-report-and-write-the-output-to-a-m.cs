using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;   // JsonDataSource resides in this namespace

class Program
{
    static void Main()
    {
        // Sample JSON data representing a list of persons.
        string json = @"{
            ""persons"": [
                { ""Name"": ""Alice"", ""Age"": 30 },
                { ""Name"": ""Bob"",   ""Age"": 25 },
                { ""Name"": ""Carol"", ""Age"": 28 }
            ]
        }";

        // Convert the JSON string to a memory stream.
        using var jsonStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json));

        // Create a JsonDataSource from the stream.
        var jsonDataSource = new JsonDataSource(jsonStream);

        // Build a simple template document with LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Template: iterate over the "persons" collection and output each person's data.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Populate the template with the JSON data.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "persons");

        // Save the generated report to a memory stream.
        using var outputStream = new MemoryStream();
        doc.Save(outputStream, SaveFormat.Docx);

        // Reset the stream position for potential further use.
        outputStream.Position = 0;

        // For demonstration, write the size of the generated document.
        Console.WriteLine($"Report generated. Size: {outputStream.Length} bytes.");
    }
}
