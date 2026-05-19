using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings that Aspose.Words might need.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Prepare sample JSON data and write it to a local file.
        // -----------------------------------------------------------------
        const string jsonFileName = "people.json";
        string jsonContent = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Carol"", ""Age"": 28 }
        ]";
        File.WriteAllText(jsonFileName, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Create a Word template that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        const string templateFileName = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templateFileName);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the JSON data source.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templateFileName);
        using (var jsonStream = File.OpenRead(jsonFileName))
        {
            var jsonDataSource = new JsonDataSource(jsonStream);
            var engine = new ReportingEngine();
            // Build the report; the root object name must match the tag reference ("persons").
            engine.BuildReport(reportDoc, jsonDataSource, "persons");
        }

        // -----------------------------------------------------------------
        // 4. Save the generated report locally.
        // -----------------------------------------------------------------
        const string reportFileName = "Report.docx";
        reportDoc.Save(reportFileName);

        // -----------------------------------------------------------------
        // 5. Simulate uploading the report to Azure Blob Storage.
        //    (Azure SDK is not part of the required packages, so we copy the file to a local folder.)
        // -----------------------------------------------------------------
        const string simulatedBlobContainer = "AzureBlobContainer";
        Directory.CreateDirectory(simulatedBlobContainer);
        string destinationPath = Path.Combine(simulatedBlobContainer, reportFileName);
        File.Copy(reportFileName, destinationPath, overwrite: true);

        // The example has completed all steps without requiring user interaction.
    }
}
