using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace MultiSectionReportExample
{
    // Simple POCO classes that will act as separate data sources.
    public class HeaderData
    {
        public string Title { get; set; }
    }

    public class BodyData
    {
        public string Content { get; set; }
    }

    public class FooterData
    {
        public string Note { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 2. Build the header section and insert a template tag that references the "Header" data source.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("<<[Header.Title]>>"); // Will be replaced by HeaderData.Title

            // 3. Build the body of the document and insert a template tag that references the "Body" data source.
            builder.MoveToDocumentEnd();
            builder.Writeln(); // Ensure we are in the main body.
            builder.Writeln("<<[Body.Content]>>"); // Will be replaced by BodyData.Content

            // 4. Build the footer section and insert a template tag that references the "Footer" data source.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln("<<[Footer.Note]>>"); // Will be replaced by FooterData.Note

            // 5. Prepare the three separate data source objects.
            var header = new HeaderData { Title = "Quarterly Report – Q1 2026" };
            var body = new BodyData
            {
                Content = "This quarter we achieved a 15% increase in revenue and launched two new products."
            };
            var footer = new FooterData { Note = "Confidential – For internal use only" };

            // 6. Use ReportingEngine to populate the template with the three data sources.
            ReportingEngine engine = new ReportingEngine
            {
                // Optional: remove any empty paragraphs that may remain after processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // The first data source name can be empty if you never need to reference the object itself,
            // but we give explicit names for clarity.
            object[] dataSources = { header, body, footer };
            string[] dataSourceNames = { "Header", "Body", "Footer" };

            engine.BuildReport(doc, dataSources, dataSourceNames);

            // 7. Save the final document.
            doc.Save("MultiSectionReport.docx");
        }
    }
}
