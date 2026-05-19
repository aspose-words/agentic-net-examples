using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReporting
{
    // Model classes
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    // Configuration class for reporting options
    public class ReportingOptionsConfig
    {
        public List<string> Options { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for files used in the example
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            string templatePath = Path.Combine(outputDir, "template.docx");
            string configPath = Path.Combine(outputDir, "reportOptions.json");
            string resultPath = Path.Combine(outputDir, "report.docx");

            // 1. Create a simple template with LINQ Reporting tags
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("People Report");
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");
            templateDoc.Save(templatePath);

            // 2. Create a JSON configuration file that defines reporting engine options
            var config = new ReportingOptionsConfig
            {
                Options = new List<string>
                {
                    "RemoveEmptyParagraphs",
                    "InlineErrorMessages"
                }
            };
            File.WriteAllText(configPath, JsonConvert.SerializeObject(config, Formatting.Indented));

            // 3. Load reporting options from the JSON configuration file
            var configJson = File.ReadAllText(configPath);
            var loadedConfig = JsonConvert.DeserializeObject<ReportingOptionsConfig>(configJson) ?? new ReportingOptionsConfig();

            ReportBuildOptions combinedOptions = ReportBuildOptions.None;
            foreach (var optionName in loadedConfig.Options)
            {
                if (Enum.TryParse(optionName, out ReportBuildOptions parsedOption))
                {
                    combinedOptions |= parsedOption;
                }
            }

            // 4. Prepare sample data for the report
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new() { Name = "Alice", Age = 30 },
                    new() { Name = "Bob", Age = 45 },
                    new() { Name = "Charlie", Age = 28 }
                }
            };

            // 5. Load the template document (must be loaded after creation)
            var doc = new Document(templatePath);

            // 6. Configure the ReportingEngine with the loaded options and build the report
            var engine = new ReportingEngine { Options = combinedOptions };
            engine.BuildReport(doc, model, "model");

            // 7. Save the generated report
            doc.Save(resultPath);
        }
    }
}
