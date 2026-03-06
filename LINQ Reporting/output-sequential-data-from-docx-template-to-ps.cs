using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsPsExport
{
    // Simple data class that will be used as the data source for the template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains Reporting Engine tags, e.g. <<[ds.Name]>> and <<[ds.Age]>>.
            const string templatePath = @"C:\Templates\PeopleReport.docx";

            // Path where the resulting PostScript file will be saved.
            const string outputPath = @"C:\Output\PeopleReport.ps";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a list of data objects that will be merged into the template.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 45 },
                new Person { Name = "Carol", Age = 27 }
            };

            // The ReportingEngine can work with any enumerable data source.
            // The data source name ("ds") must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, people, "ds");

            // Configure PostScript save options.
            PsSaveOptions psOptions = new PsSaveOptions
            {
                SaveFormat = SaveFormat.Ps, // Explicitly set the format to PostScript.
                // Additional options can be set here if needed, e.g.:
                // ColorMode = ColorMode.Normal,
                // UseHighQualityRendering = true
            };

            // Save the populated document as a PostScript file.
            doc.Save(outputPath, psOptions);
        }
    }
}
