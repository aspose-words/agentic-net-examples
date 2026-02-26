using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Simple data class used as the LINQ data source.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int    Age       { get; set; }

        // Example of an expression that can be used directly in the template.
        public string FullName => $"{FirstName} {LastName}";
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains LINQ tags, e.g. <<[ds.FullName]>>.
            const string templatePath = "Template.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create the data source instance.
            var person = new Person
            {
                FirstName = "John",
                LastName  = "Doe",
                Age       = 30
            };

            // Build the report using the LINQ Reporting Engine.
            // The data source name "ds" must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, person, "ds");

            // Configure XPS save options (optional customizations).
            XpsSaveOptions xpsOptions = new XpsSaveOptions
            {
                // Render with high‑quality algorithms.
                UseHighQualityRendering = true,
                // Example: enable output optimization.
                OptimizeOutput = true
            };

            // Save the populated document as an XPS file.
            const string outputPath = "Report.xps";
            doc.Save(outputPath, xpsOptions);
        }
    }
}
