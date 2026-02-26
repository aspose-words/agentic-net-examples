using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple class that will be used as a data source for the template.
    // The template will contain a tag like <<doc [src.Document]>> which tells the
    // ReportingEngine to insert the document referenced by this property.
    public class DocumentSource
    {
        // The document that will be inserted into the template.
        public Document Document { get; }

        public DocumentSource(string documentPath)
        {
            // Load the document that we want to embed.
            Document = new Document(documentPath);
        }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load the RTF template that contains the <<doc [src.Document]>> tag.
            // -----------------------------------------------------------------
            // The template can also contain a second tag with the "-sourceNumbering"
            // switch to keep the numbering of the inserted document separate.
            Document template = new Document(@"Templates\ReportTemplate.rtf");

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            // In this example we will insert two different documents dynamically.
            // They are wrapped in a container class so that the template can refer
            // to them via the property name.
            var firstDocSource = new DocumentSource(@"Documents\FirstPart.docx");
            var secondDocSource = new DocumentSource(@"Documents\SecondPart.docx");

            // -----------------------------------------------------------------
            // 3. Build the report.
            // -----------------------------------------------------------------
            // The ReportingEngine can accept multiple data sources.  The first
            // source name can be empty (or null) if we only need to reference its
            // members, but we give explicit names here because the template uses
            // the name "src".
            ReportingEngine engine = new ReportingEngine
            {
                // Remove any paragraphs that become empty after the tags are processed.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the two sources.  The template will replace
            // <<doc [src.Document]>> with the content of FirstPart.docx and
            // <<doc [src2.Document] -sourceNumbering>> with the content of
            // SecondPart.docx while preserving its own numbering.
            engine.BuildReport(
                template,
                new object[] { firstDocSource, secondDocSource },
                new[] { "src", "src2" });

            // -----------------------------------------------------------------
            // 4. Save the resulting document.
            // -----------------------------------------------------------------
            template.Save(@"Output\CombinedReport.docx");
        }
    }
}
