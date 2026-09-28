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
        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(workDir, "data");
        string outputDir = Path.Combine(workDir, "output");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample CSV file.
        string csvPath = Path.Combine(dataDir, "people.csv");
        File.WriteAllText(csvPath,
            "FirstName,LastName,Email\nJohn,Doe,john.doe@example.com\nJane,Smith,jane.smith@example.com",
            Encoding.UTF8);

        // Create a simple template document with LINQ Reporting tags.
        string templatePath = Path.Combine(dataDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<[record.FirstName]>> <<[record.LastName]>>");
        builder.Writeln("Email: <<[record.Email]>>");
        templateDoc.Save(templatePath);

        // Load CSV data into a list of strongly‑typed records.
        var records = new List<Record>();
        using (var reader = new StreamReader(csvPath, Encoding.UTF8))
        {
            // Read header.
            string? headerLine = reader.ReadLine();
            if (headerLine == null) return;

            // Read each data line.
            while (!reader.EndOfStream)
            {
                string? line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;

                string[] parts = line.Split(',');
                if (parts.Length < 3) continue;

                records.Add(new Record
                {
                    FirstName = parts[0],
                    LastName = parts[1],
                    Email = parts[2]
                });
            }
        }

        // Generate an individual document for each record.
        int index = 1;
        foreach (var record in records)
        {
            // Load the template.
            var doc = new Document(templatePath);

            // Build the report using the current record as the root object.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, record, "record");

            // Save the generated document.
            string outPath = Path.Combine(outputDir, $"Person_{index}.docx");
            doc.Save(outPath);
            index++;
        }
    }
}

// Public data model matching the CSV columns.
public class Record
{
    public string FirstName { get; set; } = string.Empty;
    public string LastName { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
}
