using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model for the report.
    public class ReportModel
    {
        public List<SectionData> Sections { get; set; } = new();
    }

    public class SectionData
    {
        public string Title { get; set; } = string.Empty;
        // Color can be a known color name, HTML hex code, or any value accepted by the backColor tag.
        public string Color { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Begin a foreach loop over the Sections collection.
            builder.Writeln("<<foreach [sec in Sections]>>");

            // Apply background color to the whole section using the backColor tag.
            // The color expression is taken from sec.Color.
            builder.Writeln("<<backColor [sec.Color]>>");
            // Section title.
            builder.Writeln("<<[sec.Title]>>");
            // End of backColor block.
            builder.Writeln("<</backColor>>");

            // End of foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(templatePath);

            // 2. Prepare the data source.
            var model = new ReportModel
            {
                Sections = new List<SectionData>
                {
                    new SectionData { Title = "Introduction", Color = "LightYellow" },
                    new SectionData { Title = "Details", Color = "LightBlue" },
                    new SectionData { Title = "Conclusion", Color = "#FFC0CB" } // Pink via hex code.
                }
            };

            // 3. Load the template for reporting.
            var reportDoc = new Document(templatePath);

            // 4. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 5. Save the generated report.
            reportDoc.Save("Report.docx");
        }
    }
}
