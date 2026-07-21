using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

// Wrapper that holds the original collection and provides a filtered view.
public class PersonsWrapper
{
    // Original collection – needed only for JSON deserialization.
    public List<Person> Persons { get; set; } = new();

    // Filtered collection used by the template (case‑insensitive match on "john").
    public IEnumerable<Person> FilteredPersons => Persons
        .Where(p => p.Name.Equals("john", StringComparison.OrdinalIgnoreCase));
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample data and write it to a JSON file.
        // -----------------------------------------------------------------
        var people = new List<Person>
        {
            new Person { Name = "John", Age = 30 },
            new Person { Name = "john", Age = 25 },
            new Person { Name = "Alice", Age = 28 },
            new Person { Name = "Bob", Age = 35 }
        };

        const string jsonPath = "people.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(people, Formatting.Indented));

        // -----------------------------------------------------------------
        // 2. Create a template document programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Filtered persons (case‑insensitive match on name \"john\"):");
        // Use the wrapper's FilteredPersons collection.
        builder.Writeln("<<foreach [p in FilteredPersons]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template for reporting.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Load JSON data into the wrapper object.
        // -----------------------------------------------------------------
        // Deserialize JSON into the wrapper's Persons list.
        var wrapper = new PersonsWrapper
        {
            Persons = JsonConvert.DeserializeObject<List<Person>>(File.ReadAllText(jsonPath))!
        };

        // -----------------------------------------------------------------
        // 5. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        // The root object name used in the template is "model".
        engine.BuildReport(reportDoc, wrapper, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save("output.docx");
    }
}
