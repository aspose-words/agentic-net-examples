using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath, "FirstName,LastName,Age\nJohn,Doe,30\nJane,Smith,25");

        // Load CSV records into a list of strongly‑typed objects.
        List<Record> records = LoadCsv(csvPath);

        // Create a single‑row template programmatically.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Ensure output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Generate an individual document for each record.
        for (int i = 0; i < records.Count; i++)
        {
            Record record = records[i];
            Document doc = new(templatePath);
            ReportingEngine engine = new();
            engine.BuildReport(doc, record, "record");
            string outPath = Path.Combine(outputDir, $"Record_{i + 1}.docx");
            doc.Save(outPath);
        }
    }

    // Loads CSV file into a list of Record objects.
    private static List<Record> LoadCsv(string path)
    {
        List<Record> list = new();
        string[] lines = File.ReadAllLines(path);
        if (lines.Length < 2)
            return list; // No data.

        // Assume first line contains headers.
        for (int i = 1; i < lines.Length; i++)
        {
            string line = lines[i];
            if (string.IsNullOrWhiteSpace(line))
                continue;

            string[] parts = line.Split(',');
            if (parts.Length != 3)
                continue; // Skip malformed lines.

            Record rec = new()
            {
                FirstName = parts[0],
                LastName = parts[1],
                Age = int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out int age) ? age : 0
            };
            list.Add(rec);
        }
        return list;
    }

    // Creates a Word template with LINQ Reporting tags.
    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        builder.Writeln("Report for <<[record.FirstName]>> <<[record.LastName]>>");
        builder.Writeln("Age: <<[record.Age]>>");

        doc.Save(path);
    }
}

// Public data model representing a CSV row.
public class Record
{
    public string FirstName { get; set; } = "";
    public string LastName { get; set; } = "";
    public int Age { get; set; }
}
