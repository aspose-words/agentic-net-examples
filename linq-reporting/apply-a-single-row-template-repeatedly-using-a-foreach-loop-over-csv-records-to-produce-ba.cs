using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working directory and file paths.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string csvPath = Path.Combine(workDir, "data.csv");
        string outputPath = Path.Combine(workDir, "output.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple CSV file with headers and a few rows.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Name,Age,City",
            "Alice,30,New York",
            "Bob,25,London",
            "Charlie,35,Sydney"
        };
        File.WriteAllLines(csvPath, csvLines);

        // -----------------------------------------------------------------
        // 2. Build a template document that contains a foreach tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Start the foreach block – the data source name will be "persons".
        builder.Writeln("<<foreach [person in persons]>>");
        // Inside the block output the fields from each CSV record.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("City: <<[person.City]>>");
        builder.Writeln("<</foreach>>");

        // Save the template so that the reporting engine can load it later.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template back (required before building the report).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Configure CSV data source options (has header row, comma delimiter).
        // -----------------------------------------------------------------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            HasHeaders = true
        };

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 5. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // BuildReport expects the root object name – we use "persons" to match the tag.
        engine.BuildReport(loadedTemplate, csvDataSource, "persons");

        // -----------------------------------------------------------------
        // 6. Save the generated document.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPath);
    }
}
