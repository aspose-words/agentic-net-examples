using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible CSV encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV files.
        string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataFolder);

        string csv1 = Path.Combine(dataFolder, "data1.csv");
        string csv2 = Path.Combine(dataFolder, "data2.csv");

        File.WriteAllText(csv1, "Id,Name\n1,Alpha\n2,Beta");
        File.WriteAllText(csv2, "Id,Name\n3,Gamma\n4,Delta");

        // Load and merge CSV data.
        ReportModel model = new ReportModel();
        foreach (string csvPath in new[] { csv1, csv2 })
        {
            foreach (string line in File.ReadAllLines(csvPath, Encoding.UTF8))
            {
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("Id")) continue; // Skip header.
                string[] parts = line.Split(',');
                if (parts.Length != 2) continue;
                if (int.TryParse(parts[0], out int id))
                {
                    model.Persons.Add(new Person { Id = id, Name = parts[1] });
                }
            }
        }

        // Create the template document with LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Merged CSV Report");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Id: <<[person.Id]>>, Name: <<[person.Name]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(doc, model, "model");

        // Save the result.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);
        string outputPath = Path.Combine(outputFolder, "MergedReport.docx");
        doc.Save(outputPath);
    }
}
