using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingBookmarkExample
{
    // Root data model
    public class ReportModel
    {
        public List<Section> Sections { get; set; } = new();
    }

    // Section with a bookmark name and nested subsections
    public class Section
    {
        public string Title { get; set; } = "";
        public string BookmarkName { get; set; } = "";
        public List<SubSection> SubSections { get; set; } = new();
    }

    // Subsection with its own bookmark name
    public class SubSection
    {
        public string Title { get; set; } = "";
        public string BookmarkName { get; set; } = "";
    }

    class Program
    {
        static void Main()
        {
            // Paths for template and output
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -------------------------------------------------
            // Create the template document with LINQ Reporting tags
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Title
            builder.Writeln("Report with Nested Bookmarks");
            builder.Writeln();

            // Begin outer foreach over sections
            builder.Writeln("<<foreach [section in Model.Sections]>>");

            // Outer bookmark for each section
            builder.Writeln("<<bookmark [section.BookmarkName]>>");
            builder.Writeln("Section: <<[section.Title]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln();

            // Begin inner foreach over subsections
            builder.Writeln("<<foreach [sub in section.SubSections]>>");

            // Inner bookmark for each subsection
            builder.Writeln("<<bookmark [sub.BookmarkName]>>");
            builder.Writeln(" - Subsection: <<[sub.Title]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln();

            // End inner foreach
            builder.Writeln("<</foreach>>");

            // End outer foreach
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for report generation
            // -------------------------------------------------
            var doc = new Document(templatePath);

            // -------------------------------------------------
            // Prepare sample data
            // -------------------------------------------------
            var model = new ReportModel
            {
                Sections = new List<Section>
                {
                    new Section
                    {
                        Title = "Introduction",
                        BookmarkName = "Intro",
                        SubSections = new List<SubSection>
                        {
                            new SubSection { Title = "Background", BookmarkName = "Intro_Background" },
                            new SubSection { Title = "Scope", BookmarkName = "Intro_Scope" }
                        }
                    },
                    new Section
                    {
                        Title = "Details",
                        BookmarkName = "Details",
                        SubSections = new List<SubSection>
                        {
                            new SubSection { Title = "Specifications", BookmarkName = "Details_Specs" },
                            new SubSection { Title = "Requirements", BookmarkName = "Details_Reqs" }
                        }
                    }
                }
            };

            // -------------------------------------------------
            // Build the report using LINQ Reporting Engine
            // -------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "Model");

            // Save the generated report
            doc.Save(outputPath);
        }
    }
}
