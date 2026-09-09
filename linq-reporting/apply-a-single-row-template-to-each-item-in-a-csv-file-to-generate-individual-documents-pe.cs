using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model that matches the CSV columns.
    public class Record
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public string City { get; set; } = "";
    }

    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file.
        // -----------------------------------------------------------------
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "data.csv");
        File.WriteAllText(csvPath,
            "Name,Age,City\r\n" +
            "Alice,30,New York\r\n" +
            "Bob,25,London\r\n" +
            "Charlie,35,Sydney\r\n");

        // -----------------------------------------------------------------
        // 2. Load CSV data into a list of Record objects.
        // -----------------------------------------------------------------
        List<Record> records = LoadCsv(csvPath);

        // -----------------------------------------------------------------
        // 3. Create a single‑row template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        CreateTemplate(templatePath);

        // -----------------------------------------------------------------
        // 4. For each record generate an individual document.
        // -----------------------------------------------------------------
        foreach (Record rec in records)
        {
            // Load the template fresh for each iteration.
            Document doc = new Document(templatePath);

            // Build the report using the current record as the data source.
            ReportingEngine engine = new ReportingEngine();
            // The template tags reference the root object name "record".
            engine.BuildReport(doc, rec, "record");

            // Save the generated document.
            string outPath = Path.Combine(outputDir, $"Report_{rec.Name}.docx");
            doc.Save(outPath);
        }

        Console.WriteLine("Documents generated in: " + outputDir);
    }

    // Reads a CSV file (with a header row) into a list of Record objects.
    private static List<Record> LoadCsv(string path)
    {
        var list = new List<Record>();
        string[] lines = File.ReadAllLines(path);
        if (lines.Length < 2) return list; // No data.

        // Assume first line contains headers.
        for (int i = 1; i < lines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i])) continue;
            string[] parts = lines[i].Split(',');
            if (parts.Length != 3) continue;

            var rec = new Record
            {
                Name = parts[0].Trim(),
                Age = int.TryParse(parts[1].Trim(), out int age) ? age : 0,
                City = parts[2].Trim()
            };
            list.Add(rec);
        }
        return list;
    }

    // Creates a Word template containing LINQ Reporting tags that reference a "record" object.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Name: <<[record.Name]>>");
        builder.Writeln("Age: <<[record.Age]>>");
        builder.Writeln("City: <<[record.City]>>");

        doc.Save(path);
    }
}
