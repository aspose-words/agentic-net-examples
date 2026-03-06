using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the template document that contains LINQ Reporting Engine tags.
        Document template = new Document("Template.docx");

        // Create a JSON data source from a file.
        JsonDataSource jsonData = new JsonDataSource("Data.json");

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Optional: allow missing members to be treated as null.
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report using the template and the JSON data source.
        // The third parameter is the name used to reference the data source in the template.
        engine.BuildReport(template, jsonData, "data");

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Specify a folder where images will be saved.
            ImagesFolder = "HtmlImages"
        };

        // Ensure the images folder exists.
        Directory.CreateDirectory(htmlOptions.ImagesFolder);

        // Save the populated document as an HTML file.
        template.Save("Report.html", htmlOptions);
    }
}
