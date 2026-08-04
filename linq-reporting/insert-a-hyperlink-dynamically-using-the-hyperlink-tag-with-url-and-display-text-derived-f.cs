using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

// Register code page provider (required for some Aspose.Words features)
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

namespace AsposeWordsLinqReporting
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // URL that the hyperlink will point to.
        public string Url { get; set; } = string.Empty;

        // Text that will be displayed as the hyperlink.
        public string Text { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank Word document and a builder to edit it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 2. Insert a paragraph that contains a LINQ Reporting link tag.
            //    The tag will be replaced with a hyperlink whose URL and display text
            //    come from the data model fields: Url and Text.
            builder.Writeln("Please visit: <<link [model.Url] [model.Text]>>");

            // 3. Prepare sample data.
            ReportModel model = new ReportModel
            {
                Url = "https://www.example.com",
                Text = "Example Website"
            };

            // 4. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "model", so we pass it as the third argument.
            engine.BuildReport(doc, model, "model");

            // 5. Save the resulting document.
            doc.Save("HyperlinkReport.docx");
        }
    }
}
