using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingToMarkdown
{
    // Simple data model used in the template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank Word document.
            Document doc = new Document();

            // 2. Build a template that contains LINQ Reporting tags.
            //    The tags will be replaced with the values from the data source.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Person Report");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age : <<[person.Age]>>");

            // 3. Prepare the data source.
            Person person = new Person { Name = "John Doe", Age = 42 };

            // 4. Populate the template using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter is the name used to reference the data source inside the template.
            engine.BuildReport(doc, person, "person");

            // 5. Save the populated document as Markdown.
            MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
            {
                SaveFormat = SaveFormat.Markdown // Ensure the format is Markdown.
            };
            doc.Save("PersonReport.md", mdOptions);
        }
    }
}
