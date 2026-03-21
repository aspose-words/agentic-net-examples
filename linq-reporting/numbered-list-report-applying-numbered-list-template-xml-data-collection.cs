using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

class NumberedListReport
{
    static void Main()
    {
        // Paths to the template (optional), XML data file and the output document.
        const string templatePath = "Template.docx";   // can be an empty document or a pre‑designed template
        const string xmlPath = "People.xml";
        const string outputPath = "NumberedListReport.docx";

        // -----------------------------------------------------------------
        // Ensure the XML data file exists – create a simple sample if missing.
        // -----------------------------------------------------------------
        if (!File.Exists(xmlPath))
        {
            const string sampleXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<People>
    <Person><Name>John Doe</Name></Person>
    <Person><Name>Jane Smith</Name></Person>
    <Person><Name>Bob Johnson</Name></Person>
</People>";
            File.WriteAllText(xmlPath, sampleXml);
        }

        // -----------------------------------------------------------------
        // 1. Load (or create) the Word document that will serve as the report.
        // -----------------------------------------------------------------
        Document doc = File.Exists(templatePath) ? new Document(templatePath) : new Document();

        // -----------------------------------------------------------------
        // 2. Load the XML data source.
        // -----------------------------------------------------------------
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, xmlData, "persons");

        // -----------------------------------------------------------------
        // 4. Create a numbered list based on the predefined template.
        // -----------------------------------------------------------------
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // -----------------------------------------------------------------
        // 5. Insert list items for each <Person> element in the XML.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        builder.ListFormat.List = numberedList;
        builder.ListFormat.ListLevelNumber = 0;

        XDocument xDoc = XDocument.Load(xmlPath);
        var personNames = xDoc.Descendants("Person")
                              .Select(p => (string)p.Element("Name"))
                              .Where(n => !string.IsNullOrEmpty(n));

        foreach (string name in personNames)
        {
            builder.Writeln(name);
        }

        builder.ListFormat.RemoveNumbers();

        // -----------------------------------------------------------------
        // 6. Save the final report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
