using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains LINQ Reporting tags.
        Document doc = new Document("Template.docx");

        // Prepare a sequential data source (e.g., a list of objects).
        List<Person> people = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 },
            new Person { Name = "Bob Johnson", Age = 40 }
        };

        // Populate the template with the data source using ReportingEngine.
        // The template should reference the data source by the name "people".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, people, "people");

        // Render each page of the populated document to a separate PNG file.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Select the current page for rendering.
            pngOptions.PageSet = new PageSet(pageIndex);

            // Save the page as PNG. The file name includes the page number.
            string outputFile = $"Report_Page_{pageIndex + 1}.png";
            doc.Save(outputFile, pngOptions);
        }
    }

    // Simple data class used as the LINQ Reporting data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
