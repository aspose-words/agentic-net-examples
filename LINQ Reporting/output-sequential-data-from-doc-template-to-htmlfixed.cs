using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Example data source class – replace with your own model.
    public class MyData
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the DOC template.
            Document doc = new Document("Template.docx");

            // 2. Create and populate the data source.
            MyData data = new MyData
            {
                Name = "John Doe",
                Age = 30
            };

            // 3. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter ("ds") is the name used in the template to reference the data source.
            engine.BuildReport(doc, data, "ds");

            // 4. Configure HTML Fixed save options.
            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                SaveFormat = SaveFormat.HtmlFixed,   // Explicitly set the format.
                ExportEmbeddedImages = false,        // Save images as external files.
                ResourcesFolder = "HtmlResources",   // Physical folder for resources.
                ResourcesFolderAlias = "HtmlResourcesAlias", // URI alias used in the HTML.
                ShowPageBorder = false               // Optional: hide page borders.
            };

            // Ensure the folder for external resources exists.
            Directory.CreateDirectory(saveOptions.ResourcesFolderAlias);

            // 5. Save the populated document as fixed‑layout HTML.
            doc.Save("Report.html", saveOptions);
        }
    }
}
