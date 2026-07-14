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
        // Ensure code page support for possible CSV encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a temporary folder for sample CSV files.
        string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataFolder);

        // Create two sample CSV files.
        CreateCsvFile(Path.Combine(dataFolder, "people1.csv"), new[]
        {
            "Name,Age",
            "Alice,30",
            "Bob,25"
        });

        CreateCsvFile(Path.Combine(dataFolder, "people2.csv"), new[]
        {
            "Name,Age",
            "Charlie,35",
            "Diana,28"
        });

        // Load all CSV files from the folder and merge their contents.
        ReportModel model = new();
        foreach (string csvPath in Directory.GetFiles(dataFolder, "*.csv"))
        {
            foreach (string line in File.ReadAllLines(csvPath))
            {
                // Skip header line.
                if (line.StartsWith("Name", StringComparison.OrdinalIgnoreCase))
                    continue;

                string[] parts = line.Split(',');
                if (parts.Length != 2)
                    continue;

                if (int.TryParse(parts[1], out int age))
                {
                    model.Persons.Add(new Person
                    {
                        Name = parts[0],
                        Age = age
                    });
                }
            }
        }

        // Build a template document with LINQ Reporting tags.
        Document doc = new();
        DocumentBuilder builder = new(doc);

        builder.Writeln("Report of Persons:");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Generate the report.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedReport.docx");
        doc.Save(outputPath);
    }

    private static void CreateCsvFile(string path, IEnumerable<string> lines)
    {
        File.WriteAllLines(path, lines, Encoding.UTF8);
    }
}
