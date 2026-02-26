using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReporting
{
    // DTO that will hold paragraph text and its style name.
    public class ParagraphInfo
    {
        public string Text { get; set; }
        public string StyleName { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTX template document.
            Document template = new Document("Template.dotx");

            // Project each paragraph in the template (or any source document) into the custom DTO.
            List<ParagraphInfo> data = template.GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .Select(p => new ParagraphInfo
                {
                    // Paragraph.GetText() includes a paragraph break; Trim removes it.
                    Text = p.GetText().Trim(),
                    // Retrieve the style name; fallback to "Normal" if no style is assigned.
                    StyleName = p.ParagraphFormat.Style?.Name ?? "Normal"
                })
                .ToList();

            // The template must contain a reporting region, e.g. <<foreach [paras]>>...<<[Text]>>...<<[StyleName]>>...<<end>>
            // Build the report using the collection as the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, "paras");

            // Save the populated document.
            template.Save("Result.docx");
        }
    }
}
