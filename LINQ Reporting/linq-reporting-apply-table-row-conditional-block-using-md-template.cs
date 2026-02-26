using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading; // Added for LoadOptions and LoadFormat
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = string.Empty; // Initialized to silence nullable warning
        public int Age { get; set; }
        public bool IsActive { get; set; }
    }

    public class Model
    {
        public List<Person> Persons { get; set; } = new List<Person>();
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare the markdown template.
            // The template uses LINQ Reporting syntax:
            //   <<foreach [person in Model.Persons]>>
            //   <<if [person.IsActive]>>
            //   | <<[person.Name]>> | <<[person.Age]>> |
            //   <<endif>>
            //   <<endfor>>
            string markdownTemplate = @"
| Name | Age |
|------|-----|
<<foreach [person in Model.Persons]>>
<<if [person.IsActive]>>
| <<[person.Name]>> | <<[person.Age]>> |
<<endif>>
<<endfor>>
";

            // 2. Load the markdown template into an Aspose.Words Document.
            // Use a MemoryStream to avoid creating a physical file.
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(markdownTemplate)))
            {
                var loadOptions = new LoadOptions { LoadFormat = LoadFormat.Markdown };
                Document doc = new Document(stream, loadOptions);

                // 3. Create the data source.
                var model = new Model();
                model.Persons.Add(new Person { Name = "Alice", Age = 30, IsActive = true });
                model.Persons.Add(new Person { Name = "Bob", Age = 45, IsActive = false });
                model.Persons.Add(new Person { Name = "Charlie", Age = 28, IsActive = true });

                // 4. Build the report using the ReportingEngine.
                var engine = new ReportingEngine
                {
                    // Remove empty paragraphs that may appear after processing.
                    Options = ReportBuildOptions.RemoveEmptyParagraphs
                };
                // The second overload allows us to reference the data source object itself via the name "Model".
                engine.BuildReport(doc, model, "Model");

                // 5. Save the resulting document.
                // Here we save as DOCX, but you could also save back to Markdown or any other supported format.
                doc.Save("ReportResult.docx", SaveFormat.Docx);
            }

            Console.WriteLine("Report generated successfully.");
        }
    }
}
