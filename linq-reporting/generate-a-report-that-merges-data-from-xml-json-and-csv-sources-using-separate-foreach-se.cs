using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for older code pages).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create sample data files (XML, JSON, CSV) in the working directory.
        // -----------------------------------------------------------------
        string xmlPath = "people.xml";
        File.WriteAllText(xmlPath,
@"<People>
    <Person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </Person>
    <Person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </Person>
</People>");

        string jsonPath = "people.json";
        File.WriteAllText(jsonPath,
@"[
    { ""Name"": ""Alice Brown"", ""Age"": 28 },
    { ""Name"": ""Bob Johnson"", ""Age"": 35 }
]");

        string csvPath = "people.csv";
        File.WriteAllText(csvPath,
@"Name,Age
Charlie Davis,40
Diana Evans,22");

        // --------------------------------------------------------------
        // Build the template document programmatically (required by the rules).
        // --------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // XML data section.
        builder.Writeln("XML Data:");
        builder.Writeln("<<foreach [p in xml]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // JSON data section.
        builder.Writeln("JSON Data:");
        builder.Writeln("<<foreach [j in json]>>");
        builder.Writeln("- <<[j.Name]>> (Age: <<[j.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // CSV data section.
        builder.Writeln("CSV Data:");
        builder.Writeln("<<foreach [c in csv]>>");
        builder.Writeln("- <<[c.Name]>> (Age: <<[c.Age]>>)");
        builder.Writeln("<</foreach>>");

        // Save the template and reload it to satisfy the lifecycle rule.
        string templatePath = "template.docx";
        template.Save(templatePath);
        Document doc = new Document(templatePath);

        // --------------------------------------------------------------
        // Create data source objects.
        // --------------------------------------------------------------
        var xmlData = new XmlDataSource(xmlPath);
        var jsonData = new JsonDataSource(jsonPath);

        // Parse CSV into a strongly‑typed list so that LINQ Reporting can access
        // members via property names (avoids DataRow member errors).
        List<Person> csvData = LoadCsv(csvPath);

        // --------------------------------------------------------------
        // Build the report using multiple data sources.
        // --------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc,
            new object[] { xmlData, jsonData, csvData },
            new string[] { "xml", "json", "csv" });

        // Save the final report.
        doc.Save("Report.docx");
    }

    // Simple model class used for CSV data.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Reads a CSV file with a header line and returns a list of Person objects.
    private static List<Person> LoadCsv(string path)
    {
        var people = new List<Person>();
        using (var reader = new StreamReader(path))
        {
            // Read header line.
            string? headerLine = reader.ReadLine();
            if (headerLine == null)
                return people; // Empty file.

            // Expecting "Name,Age" – split on commas.
            while (!reader.EndOfStream)
            {
                string? line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] parts = line.Split(',');
                if (parts.Length < 2)
                    continue;

                var person = new Person
                {
                    Name = parts[0].Trim(),
                    Age = int.TryParse(parts[1].Trim(), out int age) ? age : 0
                };
                people.Add(person);
            }
        }
        return people;
    }
}
