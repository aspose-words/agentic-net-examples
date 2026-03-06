using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // This class demonstrates loading a DOC template, populating it with data,
    // and saving the result as a PostScript (PS) file.
    public class DocToPsConverter
    {
        // Replace with the actual path to your template document.
        private const string TemplatePath = @"C:\Templates\ReportTemplate.docx";

        // Replace with the desired output path for the PS file.
        private const string OutputPsPath = @"C:\Output\Report.ps";

        // Example data source class. In a real scenario, replace with your own data model.
        public class ReportData
        {
            public string Title { get; set; }
            public string Author { get; set; }
            public DateTime CreatedDate { get; set; }
            // Add additional properties as needed for the template.
        }

        public static void Main()
        {
            // 1. Load the DOC template.
            Document doc = new Document(TemplatePath);

            // 2. Populate the template using Aspose.Words.ReportingEngine.
            //    The template should contain tags like <<[data.Title]>> etc.
            var data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Author = "John Doe",
                CreatedDate = DateTime.Now
            };

            ReportingEngine engine = new ReportingEngine();
            // BuildReport returns a bool indicating success; we ignore it here.
            engine.BuildReport(doc, data, "data");

            // 3. Configure PS save options if needed.
            PsSaveOptions psOptions = new PsSaveOptions
            {
                // Example: embed generator name (default true) and set color mode.
                ExportGeneratorName = true,
                ColorMode = ColorMode.Normal,
                // Ensure the format is set to PS (redundant but explicit).
                SaveFormat = SaveFormat.Ps
            };

            // 4. Save the populated document as a PostScript file.
            doc.Save(OutputPsPath, psOptions);

            Console.WriteLine("Document has been successfully saved to PS format at:");
            Console.WriteLine(OutputPsPath);
        }
    }
}
