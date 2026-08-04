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
    public string City { get; set; } = "";
}

public class ReportModel
{
    public List<Person> People { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support (kept for completeness)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string workingDir = Directory.GetCurrentDirectory();
        string csvPath = Path.Combine(workingDir, "data.csv");
        string templatePath = Path.Combine(workingDir, "Template.docx");
        string reportPath = Path.Combine(workingDir, "Report.docx");

        // Create sample CSV file
        File.WriteAllText(csvPath,
            "Name,Age,City\r\n" +
            "Alice,30,New York\r\n" +
            "Bob,25,Los Angeles\r\n" +
            "Charlie,35,Chicago");

        // Load CSV data into strongly‑typed model
        ReportModel model = new ReportModel();
        foreach (var line in File.ReadAllLines(csvPath))
        {
            // Skip header
            if (line.StartsWith("Name,"))
                continue;

            var parts = line.Split(',');
            if (parts.Length != 3)
                continue;

            model.People.Add(new Person
            {
                Name = parts[0],
                Age = int.TryParse(parts[1], out var age) ? age : 0,
                City = parts[2]
            });
        }

        // Create a DOCX template with custom margins
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Set custom margins (2 cm on each side)
        const double cmToPoints = 28.3465;
        float margin = (float)(2 * cmToPoints);
        builder.PageSetup.TopMargin = margin;
        builder.PageSetup.BottomMargin = margin;
        builder.PageSetup.LeftMargin = margin;
        builder.PageSetup.RightMargin = margin;

        // Add a title
        builder.Writeln("People Report");
        builder.Writeln();

        // Insert LINQ Reporting tags to iterate over the People collection
        builder.Writeln("<<foreach [person in People]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>, City: <<[person.City]>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for report generation
        Document reportDoc = new Document(templatePath);

        // Build the report using the strongly‑typed model
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report
        reportDoc.Save(reportPath);
    }
}
