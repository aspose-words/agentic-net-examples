using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Simple data class used as the data source for the template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains LINQ Reporting tags.
            // Example tag in the template: <<foreach [people]>><<[Name]>> (Age: <<[Age]>>)<<</foreach>>
            string templatePath = @"C:\Templates\PeopleReport.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a list of data objects.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 45 },
                new Person { Name = "Carol", Age = 27 }
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "people" must match the name used in the template tags.
            engine.BuildReport(doc, people, "people");

            // Save the populated document as Markdown.
            // Using MarkdownSaveOptions allows fine‑tuned control over the output.
            MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
            {
                SaveFormat = SaveFormat.Markdown,
                ExportImagesAsBase64 = false, // Images will be saved as separate files.
                ImagesFolder = @"C:\Output\Images", // Folder for extracted images.
                ImagesFolderAlias = "images" // URI prefix used in the Markdown file.
            };

            string outputPath = @"C:\Output\PeopleReport.md";
            doc.Save(outputPath, mdOptions);
        }
    }
}
