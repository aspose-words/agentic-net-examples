using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Configuration model for reporting engine options.
    public class ReportConfig
    {
        public List<string> Options { get; set; } = new();
    }

    // Simple data model used in the JSON data source.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public static void Main()
    {
        // Ensure code page support (required by Aspose.Words for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Paths for files used in the example.
        const string configPath = "reportConfig.json";
        const string dataPath = "data.json";
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // 1. Create a JSON configuration file that defines ReportingEngine options.
        var config = new ReportConfig
        {
            Options = new() { "AllowMissingMembers", "RemoveEmptyParagraphs" }
        };
        File.WriteAllText(configPath, JsonConvert.SerializeObject(config, Formatting.Indented));

        // 2. Create a JSON data file that will be used as the data source.
        var people = new List<Person>
        {
            new() { Name = "Alice", Age = 30 },
            new() { Name = "Bob", Age = 45 },
            new() { Name = "Charlie", Age = 28 }
        };
        File.WriteAllText(dataPath, JsonConvert.SerializeObject(people, Formatting.Indented));

        // 3. Programmatically build a template document containing LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 4. Load the template document.
        var doc = new Document(templatePath);

        // 5. Load reporting options from the external JSON configuration file.
        var configJson = File.ReadAllText(configPath);
        var loadedConfig = JsonConvert.DeserializeObject<ReportConfig>(configJson) ?? new();

        // 6. Translate option strings into ReportBuildOptions flags.
        var engineOptions = ReportBuildOptions.None;
        foreach (var opt in loadedConfig.Options)
        {
            engineOptions |= opt switch
            {
                "AllowMissingMembers" => ReportBuildOptions.AllowMissingMembers,
                "RemoveEmptyParagraphs" => ReportBuildOptions.RemoveEmptyParagraphs,
                "InlineErrorMessages" => ReportBuildOptions.InlineErrorMessages,
                "UseLegacyHeaderFooterVisiting" => ReportBuildOptions.UseLegacyHeaderFooterVisiting,
                "RespectJpegExifOrientation" => ReportBuildOptions.RespectJpegExifOrientation,
                "UpdateFieldsSyntaxAware" => ReportBuildOptions.UpdateFieldsSyntaxAware,
                _ => ReportBuildOptions.None
            };
        }

        // 7. Create the reporting engine and apply the loaded options.
        var engine = new ReportingEngine { Options = engineOptions };

        // 8. Create a JsonDataSource from the data file.
        var dataSource = new JsonDataSource(dataPath);

        // 9. Build the report. The data source name "persons" matches the tag in the template.
        engine.BuildReport(doc, dataSource, "persons");

        // 10. Save the generated report.
        doc.Save(outputPath);
    }
}
