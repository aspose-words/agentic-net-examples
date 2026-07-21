using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data entity representing a person.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Person> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create sample XML data.
            // -----------------------------------------------------------------
            const string xmlFile = "people.xml";
            var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<People>
    <Person><Name>John Doe</Name><Age>28</Age></Person>
    <Person><Name>Jane Smith</Name><Age>35</Age></Person>
    <Person><Name>Bob Johnson</Name><Age>42</Age></Person>
    <Person><Name>Alice Brown</Name><Age>31</Age></Person>
</People>";
            File.WriteAllText(xmlFile, xmlContent);

            // -----------------------------------------------------------------
            // 2. Load XML, filter nodes (Age > 30), and map to model objects.
            // -----------------------------------------------------------------
            XDocument xDoc = XDocument.Load(xmlFile);
            var filteredPersons = xDoc.Root?
                .Elements("Person")
                .Where(p => (int)p.Element("Age")! > 30)
                .Select(p => new Person
                {
                    Name = (string)p.Element("Name")!,
                    Age = (int)p.Element("Age")!
                })
                .ToList() ?? new List<Person>();

            var model = new ReportModel { Items = filteredPersons };

            // -----------------------------------------------------------------
            // 3. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            const string templatePath = "template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People Report (Age > 30)");
            // The <<restartNum>> tag ensures numbering starts at 1 for this list.
            // The foreach tag iterates over the Items collection in the model.
            builder.Writeln("<<restartNum>><<foreach [person in Items]>><<[person.Name]>> (Age: <<[person.Age]>>)"); 
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 4. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after tag processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "PeopleReport.docx";
            doc.Save(outputPath);

            // The example finishes without waiting for user input.
        }
    }
}
