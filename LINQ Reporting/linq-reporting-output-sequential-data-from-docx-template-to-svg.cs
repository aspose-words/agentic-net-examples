using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingToSvg
{
    // Simple data model that will be used as the data source for the LINQ Reporting Engine.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the DOCX template that contains LINQ Reporting tags, e.g. <<foreach [persons]>><<[Name]>> (<[Age]>)<</foreach>>.
            Document doc = new Document("Template.docx");

            // Prepare sequential data.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 25 },
                new Person { Name = "Carol", Age = 28 }
            };

            // Populate the template with the data using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter is the name by which the data source is referenced inside the template.
            engine.BuildReport(doc, people, "persons");

            // Configure SVG save options (optional – here we render text as placed glyphs).
            SvgSaveOptions svgOptions = new SvgSaveOptions
            {
                TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
                // Fit the whole document into the SVG viewport.
                FitToViewPort = true,
                // Hide page borders for a cleaner SVG.
                ShowPageBorder = false
            };

            // Save the populated document as an SVG file.
            doc.Save("Report.svg", svgOptions);
        }
    }
}
