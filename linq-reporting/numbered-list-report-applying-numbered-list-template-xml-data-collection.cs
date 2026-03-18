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
        // 1. Load (or create) the Word document that will serve as the report.
        // -----------------------------------------------------------------
        Document doc;
        if (File.Exists(templatePath))
        {
            // Load an existing template if it exists.
            doc = new Document(templatePath);
        }
        else
        {
            // Otherwise create a new blank document.
            doc = new Document();
        }

        // -----------------------------------------------------------------
        // 2. Load the XML data source.
        // -----------------------------------------------------------------
        // The XML file is expected to have a simple structure, e.g.:
        // <People>
        //   <Person><Name>John Doe</Name></Person>
        //   <Person><Name>Jane Smith</Name></Person>
        //   ...
        // </People>
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        //    The data source name ("persons") will be used inside the template
        //    if the template contains any <<[persons]>> tags.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The BuildReport overload with a data source name is used so that the
        // template can reference the source directly (e.g. <<foreach [persons.Person]>>).
        engine.BuildReport(doc, xmlData, "persons");

        // -----------------------------------------------------------------
        // 4. Create a numbered list based on the predefined template.
        // -----------------------------------------------------------------
        // ListTemplate.NumberDefault corresponds to the default multi‑level numbered list.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // -----------------------------------------------------------------
        // 5. Insert list items for each <Person> element in the XML.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Position the builder at the end of the document (after any template content).
        builder.MoveToDocumentEnd();

        // Apply the list to the builder once; subsequent paragraphs will inherit it.
        builder.ListFormat.List = numberedList;
        // Ensure we start at the first level (0‑based index).
        builder.ListFormat.ListLevelNumber = 0;

        // Load the XML into LINQ for easy enumeration.
        XDocument xDoc = XDocument.Load(xmlPath);
        var personNames = xDoc.Descendants("Person")
                              .Select(p => (string)p.Element("Name"))
                              .Where(n => !string.IsNullOrEmpty(n));

        foreach (string name in personNames)
        {
            // Each WriteLine creates a new paragraph that automatically receives the list formatting.
            builder.Writeln(name);
        }

        // Remove list formatting from the builder so that any following content is normal text.
        builder.ListFormat.RemoveNumbers();

        // -----------------------------------------------------------------
        // 6. Save the final report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
