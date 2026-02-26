using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingToSvg
{
    // Simple data model for LINQ reporting.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the PDF template that contains LINQ Reporting tags (e.g. <<foreach [persons]>><<[Name]>> <<[Age]>> <<[City]>> <</foreach>>).
            Document template = new Document("Template.pdf");

            // Prepare a sequential data source using LINQ (a List of Person objects).
            List<Person> persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, City = "New York" },
                new Person { Name = "Bob",   Age = 25, City = "London"   },
                new Person { Name = "Carol", Age = 28, City = "Paris"    }
            };

            // Build the report by populating the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "persons" must match the name used in the template tags.
            engine.BuildReport(template, persons, "persons");

            // Configure SVG save options.
            SvgSaveOptions svgOptions = new SvgSaveOptions
            {
                // Render text as placed glyphs so the output looks identical to the PDF.
                TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
                // Optional: remove page borders for a cleaner SVG.
                ShowPageBorder = false,
                // Optional: fit the SVG to the viewport.
                FitToViewPort = true
            };

            // Save the populated document as an SVG file.
            template.Save("ReportOutput.svg", svgOptions);
        }
    }
}
