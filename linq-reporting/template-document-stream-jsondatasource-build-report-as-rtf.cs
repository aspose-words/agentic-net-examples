using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class ReportGenerator
{
    /// <summary>
    /// Loads a Word template from a stream, populates it with JSON data from a stream,
    /// and saves the resulting document as an RTF file.
    /// </summary>
    /// <param name="templateStream">Stream containing the template .docx/.doc file.</param>
    /// <param name="jsonStream">Stream containing the JSON data source.</param>
    /// <param name="outputPath">Full file path where the RTF document will be saved.</param>
    public void GenerateReport(Stream templateStream, Stream jsonStream, string outputPath)
    {
        // Load the template document from the provided stream.
        Document template = new Document(templateStream);

        // Read the JSON text from the stream – Aspose.Words.Reporting.JsonDataSource expects a string.
        string jsonText;
        using (var reader = new StreamReader(jsonStream))
        {
            jsonText = reader.ReadToEnd();
        }
        JsonDataSource jsonDataSource = new JsonDataSource(jsonText);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The third parameter is the name used to reference the data source in the template.
        engine.BuildReport(template, jsonDataSource, "data");

        // Prepare RTF save options (default options are sufficient for most cases).
        RtfSaveOptions rtfOptions = new RtfSaveOptions();

        // Save the populated document as RTF.
        template.Save(outputPath, rtfOptions);
    }

    // Entry point required for a console application.
    public static void Main(string[] args)
    {
        if (args.Length != 3)
        {
            Console.WriteLine("Usage: ReportGenerator <templatePath> <jsonPath> <outputRtfPath>");
            return;
        }

        string templatePath = args[0];
        string jsonPath = args[1];
        string outputPath = args[2];

        using (FileStream templateStream = File.OpenRead(templatePath))
        using (FileStream jsonStream = File.OpenRead(jsonPath))
        {
            var generator = new ReportGenerator();
            generator.GenerateReport(templateStream, jsonStream, outputPath);
        }

        Console.WriteLine($"Report generated successfully: {outputPath}");
    }
}
