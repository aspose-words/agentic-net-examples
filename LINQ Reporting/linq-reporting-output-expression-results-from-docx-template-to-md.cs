using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqToMarkdown
{
    // Simple data model used by the LINQ reporting engine.
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
            //    The tags are enclosed in << >> and can reference members of the data source.
            //    Here we use a foreach loop to list all persons.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("People Report");
            builder.Writeln("==============");
            builder.Writeln(""); // empty line for readability
            builder.Writeln("<<foreach [data]>><<[Name]>> (Age: <<[Age]>>)");
            builder.Writeln("<</foreach>>");

            // 3. Prepare the data source – a list of Person objects.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 45 },
                new Person { Name = "Carol", Age = 27 }
            };

            // 4. Populate the template using the ReportingEngine.
            //    The third argument ("data") is the name used inside the template to reference the source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, people, "data");

            // 5. Save the populated document as Markdown.
            MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
            {
                SaveFormat = SaveFormat.Markdown   // enforce Markdown format
            };
            doc.Save("PeopleReport.md", mdOptions);
        }
    }
}
