using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReportingInternalLinks
{
    // Data model classes
    public class ReportModel
    {
        public List<Section> Sections { get; set; } = new();
    }

    public class Section
    {
        public string Title { get; set; } = "";
        public string BookmarkName { get; set; } = "";
        public string Content { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output documents.
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title for the document.
            builder.Writeln("Document with Internal Links");
            builder.Writeln();

            // Table of Contents (list of links to sections).
            builder.Writeln("Table of Contents:");
            builder.Writeln("<<foreach [sec in model.Sections]>>");
            // Link tag: first expression is bookmark name, second is display text.
            builder.Writeln("<<link [sec.BookmarkName] [sec.Title]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln();

            // Sections with bookmarks.
            builder.Writeln("<<foreach [sec in model.Sections]>>");
            // Bookmark start tag.
            builder.Writeln("<<bookmark [sec.BookmarkName]>>");
            // Section title.
            builder.Writeln("<<[sec.Title]>>");
            // End of bookmark.
            builder.Writeln("<</bookmark>>");
            // Section content.
            builder.Writeln("<<[sec.Content]>>");
            builder.Writeln();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Prepare sample data.
            // -------------------------------------------------
            var model = new ReportModel
            {
                Sections = new List<Section>
                {
                    new Section
                    {
                        Title = "Introduction",
                        BookmarkName = "bmIntro",
                        Content = "This is the introduction section."
                    },
                    new Section
                    {
                        Title = "Details",
                        BookmarkName = "bmDetails",
                        Content = "Here are the detailed contents of the report."
                    },
                    new Section
                    {
                        Title = "Conclusion",
                        BookmarkName = "bmConclusion",
                        Content = "The report concludes with these remarks."
                    }
                }
            };

            // -------------------------------------------------
            // 3. Load the template and build the report.
            // -------------------------------------------------
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 4. Save the generated document.
            // -------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
