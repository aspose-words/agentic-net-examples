using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Enable reflection optimization globally.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = true;

        // -----------------------------------------------------------------
        // 2. Build a report using a regular object data source.
        // -----------------------------------------------------------------
        // Create sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var docObject = new Document();
        var builderObject = new DocumentBuilder(docObject);
        builderObject.Writeln("<<foreach [p in Persons]>>");
        builderObject.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builderObject.Writeln("<</foreach>>");

        // Build the report.
        var engineObject = new ReportingEngine();
        engineObject.BuildReport(docObject, model, "model");

        // Save the result.
        docObject.Save("ObjectReport.docx");

        // -----------------------------------------------------------------
        // 3. Disable reflection optimization for a small CSV import.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = false;

        // Prepare a small CSV file.
        string csvPath = "people.csv";
        File.WriteAllText(csvPath,
            "Name,Age\n" +
            "David,40\n" +
            "Eva,28\n" +
            "Frank,33");

        // Create a template for the CSV data source.
        var docCsv = new Document();
        var builderCsv = new DocumentBuilder(docCsv);
        builderCsv.Writeln("<<foreach [row in persons]>>");
        builderCsv.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builderCsv.Writeln("<</foreach>>");

        // Load CSV data with headers.
        var loadOptions = new CsvDataLoadOptions(true);
        var csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the CSV data source.
        var engineCsv = new ReportingEngine();
        engineCsv.BuildReport(docCsv, csvDataSource, "persons");

        // Save the CSV‑based report.
        docCsv.Save("CsvReport.docx");
    }
}

// ---------------------------------------------------------------------
// Data model for the object‑based report.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
