using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingAlignmentExample
{
    // Data model classes
    public class ReportItem
    {
        public int Index { get; set; } = 0;
        public string Text { get; set; } = "";
        public string Alignment { get; set; } = "Left"; // Left, Center, Right, Justify
    }

    public class ReportModel
    {
        public List<ReportItem> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data
            var model = new ReportModel
            {
                Items = new List<ReportItem>
                {
                    new ReportItem { Index = 0, Text = "First paragraph, left aligned.", Alignment = "Left" },
                    new ReportItem { Index = 1, Text = "Second paragraph, centered.", Alignment = "Center" },
                    new ReportItem { Index = 2, Text = "Third paragraph, right aligned.", Alignment = "Right" }
                }
            };

            // Create template document programmatically
            const string templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a foreach loop with a bookmark for each item
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("<<bookmark [\"bm_\" + item.Index]>>");
            builder.Writeln("<<[item.Text]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<</foreach>>");

            // Save the template
            doc.Save(templatePath);

            // Load the template for reporting
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report
            engine.BuildReport(reportDoc, model, "model");

            // After report generation, adjust paragraph alignment based on each item's Alignment property
            foreach (var item in model.Items)
            {
                string bookmarkName = $"bm_{item.Index}";
                var bookmark = reportDoc.Range.Bookmarks[bookmarkName];
                if (bookmark != null && bookmark.BookmarkStart?.ParentNode is Paragraph paragraph)
                {
                    paragraph.ParagraphFormat.Alignment = GetParagraphAlignment(item.Alignment);
                }
            }

            // Save the final document
            const string outputPath = "output.docx";
            reportDoc.Save(outputPath);
        }

        // Helper to map string alignment to Aspose.Words ParagraphAlignment
        private static ParagraphAlignment GetParagraphAlignment(string alignment)
        {
            return alignment?.ToLower() switch
            {
                "center" => ParagraphAlignment.Center,
                "right" => ParagraphAlignment.Right,
                "justify" => ParagraphAlignment.Justify,
                _ => ParagraphAlignment.Left,
            };
        }
    }
}
