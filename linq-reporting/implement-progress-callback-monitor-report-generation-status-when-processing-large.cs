using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace ReportGenerationWithProgress
{
    // Callback that receives loading progress notifications.
    public class LoadingProgressCallback : IDocumentLoadingCallback
    {
        public void Notify(DocumentLoadingArgs args)
        {
            // Output the estimated loading progress percentage.
            Console.WriteLine($"Loading progress: {args.EstimatedProgress}%");
        }
    }

    // Callback that receives saving progress notifications.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        public void Notify(DocumentSavingArgs args)
        {
            // Output the estimated saving progress percentage.
            Console.WriteLine($"Saving progress: {args.EstimatedProgress}%");
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the template document that contains reporting tags.
            string templatePath = @"C:\Data\ReportTemplate.docx";

            // Path to the large XML data file.
            string xmlDataPath = @"C:\Data\LargeDataSet.xml";

            // Load the template document with a loading progress callback.
            LoadOptions loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback()
            };
            Document template = new Document(templatePath, loadOptions);

            // Load the XML data source (no LoadOptions are required for XmlDataSource).
            XmlDataSource dataSource = new XmlDataSource(xmlDataPath);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "Data");

            // Save the generated report with a saving progress callback.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback()
            };
            string outputPath = @"C:\Data\GeneratedReport.docx";
            template.Save(outputPath, saveOptions);

            Console.WriteLine("Report generation completed.");
        }
    }
}
