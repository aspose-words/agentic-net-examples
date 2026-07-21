using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        List<Person> people = new()
        {
            new Person("John Doe", 30, "john.doe@example.com"),
            new Person("Jane Smith", 25, "jane.smith@example.com"),
            new Person("Bob Johnson", 45, "bob.johnson@example.com")
        };

        // Serialize data to JSON and write to a local file.
        string jsonPath = "people.json";
        string jsonContent = JsonConvert.SerializeObject(people, Formatting.Indented);
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // Create a JSON data source for the reporting engine.
        JsonDataSource jsonDataSource = new(jsonPath);

        // Build the template document programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);

        // Begin a foreach loop over the JSON array (named "persons").
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name : <<[person.Name]>>");
        builder.Writeln("Age  : <<[person.Age]>>");
        builder.Writeln("Email: <<[person.Email]>>");
        // Insert a page break after each record for separate sections.
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("<</foreach>>");

        // Generate the report using the reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(template, jsonDataSource, "persons");

        // Save the final document.
        template.Save("Report.docx");
    }
}

// Public data model class matching the JSON structure.
public class Person
{
    public Person(string name, int age, string email)
    {
        Name = name;
        Age = age;
        Email = email;
    }

    public string Name { get; set; }
    public int Age { get; set; }
    public string Email { get; set; }
}
