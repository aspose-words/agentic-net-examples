using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        const string jsonFilePath = "people.json";
        const string reportFilePath = "PeopleReport.docx";

        // 1. Create sample JSON data.
        var people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob", Age = 45 },
            new Person { Name = "Charlie", Age = 28 }
        };
        string jsonContent = JsonSerializer.Serialize(people);
        File.WriteAllText(jsonFilePath, jsonContent);

        // 2. Build a Word template with LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // 3. Load JSON data from a file stream.
        using (FileStream jsonStream = File.OpenRead(jsonFilePath))
        {
            JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(doc, jsonDataSource, "persons");
        }

        // 5. Save the generated report locally.
        doc.Save(reportFilePath);

        // 6. Simulate uploading the report to Azure Blob Storage.
        // In a real scenario you would use Azure.Storage.Blobs SDK.
        // Here we simply copy the file to a local folder named "blobstorage".
        const string simulatedContainer = "blobstorage";
        Directory.CreateDirectory(simulatedContainer);
        string destinationPath = Path.Combine(simulatedContainer, Path.GetFileName(reportFilePath));
        File.Copy(reportFilePath, destinationPath, overwrite: true);
        Console.WriteLine($"Report copied to simulated blob container: {destinationPath}");

        // Cleanup temporary files (optional).
        File.Delete(jsonFilePath);
        File.Delete(reportFilePath);
    }

    // Simple data model matching the JSON structure.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}
