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
    public string Address { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for .NET Core (required by Aspose.Words for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data (array of Person objects).
        string jsonPath = "people.json";
        var people = new List<Person>
        {
            new Person { Name = "Alice Johnson", Age = 30, Address = "123 Maple St, Springfield" },
            new Person { Name = "Bob Smith", Age = 45, Address = "456 Oak Ave, Metropolis" },
            new Person { Name = "Carol Davis", Age = 28, Address = "789 Pine Rd, Smalltown" }
        };
        string jsonContent = System.Text.Json.JsonSerializer.Serialize(people, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // Step 1: Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title for the report.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin foreach loop over the JSON array named "persons".
        builder.Writeln("<<foreach [person in persons]>>");

        // Each person gets a separate section with a heading and details.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Address: <<[person.Address]>>");
        builder.Writeln(); // Blank line between entries.

        // End of foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report using JSON data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Create a JsonDataSource from the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report. The root object name in the template is "persons".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // -----------------------------------------------------------------
        // Step 3: Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = "PeopleReport.docx";
        reportDoc.Save(outputPath);

        // Inform that the process completed.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
