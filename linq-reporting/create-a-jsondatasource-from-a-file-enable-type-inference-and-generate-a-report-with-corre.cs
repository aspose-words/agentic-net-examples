using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create a sample JSON file with a collection of people.
        string jsonPath = Path.Combine(workDir, "people.json");
        string jsonContent = @"
[
  {
    ""Name"": ""Alice Johnson"",
    ""Age"": 30,
    ""JoinDate"": ""2022-05-01T00:00:00"",
    ""IsMember"": true
  },
  {
    ""Name"": ""Bob Smith"",
    ""Age"": 45,
    ""JoinDate"": ""2020-11-15T00:00:00"",
    ""IsMember"": false
  },
  {
    ""Name"": ""Carol White"",
    ""Age"": 27,
    ""JoinDate"": ""2023-01-20T00:00:00"",
    ""IsMember"": true
  }
]";
        File.WriteAllText(jsonPath, jsonContent.Trim());

        // 2. Build a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("People Report");
        builder.Writeln("==============");
        // Begin a foreach loop over the JSON array named 'persons'.
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name      : <<[p.Name]>>");
        builder.Writeln("Age       : <<[p.Age]>>");
        builder.Writeln("Join Date : <<[p.JoinDate]>>");
        builder.Writeln("Member    : <<[p.IsMember]>>");
        builder.Writeln("------------------------------");
        // End of the foreach block.
        builder.Writeln("<</foreach>>");

        // 3. Configure JSON data load options (type inference is enabled by default).
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            // The following settings illustrate typical usage; they are optional.
            SimpleValueParseMode = JsonSimpleValueParseMode.Loose,
            PreserveSpaces = false,
            AlwaysGenerateRootObject = false
        };

        // 4. Create a JsonDataSource from the file with the specified options.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, loadOptions);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            // No special options are required for this simple example.
            Options = ReportBuildOptions.None
        };
        // The data source name must match the name used in the template tags ('persons').
        engine.BuildReport(doc, jsonDataSource, "persons");

        // 6. Save the generated report.
        string outputPath = Path.Combine(workDir, "PeopleReport.docx");
        doc.Save(outputPath);

        // The program finishes without waiting for user input.
    }
}
