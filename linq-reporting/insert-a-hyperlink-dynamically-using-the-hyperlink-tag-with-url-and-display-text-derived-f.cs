using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model used by the LINQ Reporting engine.
    public class LinkInfo
    {
        // URL that the hyperlink will point to.
        public string Url { get; set; } = "https://www.example.com";

        // Text that will be displayed for the hyperlink.
        public string Text { get; set; } = "Visit Example";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a template document that contains a LINQ Reporting link tag.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // The tag uses the model name "model" and references its Url and Text fields.
            builder.Writeln("<<link [model.Url] [model.Text]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report using a data source.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Prepare sample data.
            LinkInfo data = new LinkInfo
            {
                Url = "https://www.aspose.com",
                Text = "Aspose.Words Home"
            };

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, data, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
