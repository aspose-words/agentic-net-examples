using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingBatch
{
    // Simple data model representing a person.
    public class Person
    {
        public int Id { get; set; } = 0;
        public string Name { get; set; } = string.Empty;
    }

    // Wrapper model that will be passed as the root object to the reporting engine.
    public class BatchReport
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV encoding support.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare a temporary folder for sample CSV files.
            string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "BatchData");
            Directory.CreateDirectory(dataFolder);

            // Create sample CSV files.
            CreateSampleCsv(Path.Combine(dataFolder, "data1.csv"), new[]
            {
                new Person { Id = 1, Name = "Alice" },
                new Person { Id = 2, Name = "Bob" }
            });

            CreateSampleCsv(Path.Combine(dataFolder, "data2.csv"), new[]
            {
                new Person { Id = 3, Name = "Charlie" },
                new Person { Id = 4, Name = "Diana" }
            });

            // Load all CSV files and merge their contents into a single list.
            var reportModel = new BatchReport();

            foreach (string csvFile in Directory.GetFiles(dataFolder, "*.csv"))
            {
                foreach (var person in LoadCsv(csvFile))
                {
                    reportModel.Persons.Add(person);
                }
            }

            // Create the LINQ Reporting template programmatically.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            CreateTemplate(templatePath);

            // Load the template document.
            Document doc = new Document(templatePath);

            // Build the report using the merged data.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, reportModel, "model");

            // Save the generated report.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BatchReportOutput.docx");
            doc.Save(outputPath);
        }

        // Writes a CSV file with a header and the supplied person records.
        private static void CreateSampleCsv(string filePath, IEnumerable<Person> persons)
        {
            using var writer = new StreamWriter(filePath, false, Encoding.UTF8);
            writer.WriteLine("Id,Name");
            foreach (var p in persons)
            {
                writer.WriteLine($"{p.Id},{p.Name}");
            }
        }

        // Reads a CSV file and returns a list of Person objects.
        private static List<Person> LoadCsv(string filePath)
        {
            var persons = new List<Person>();
            using var reader = new StreamReader(filePath, Encoding.UTF8);
            bool isFirstLine = true;
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                if (isFirstLine)
                {
                    // Skip header.
                    isFirstLine = false;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] parts = line.Split(',');
                if (parts.Length >= 2 &&
                    int.TryParse(parts[0], out int id))
                {
                    persons.Add(new Person
                    {
                        Id = id,
                        Name = parts[1]
                    });
                }
            }
            return persons;
        }

        // Generates a simple Word template containing a foreach loop over the Persons collection.
        private static void CreateTemplate(string templatePath)
        {
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Batch Report");
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Id: <<[p.Id]>>, Name: <<[p.Name]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);
        }
    }
}
