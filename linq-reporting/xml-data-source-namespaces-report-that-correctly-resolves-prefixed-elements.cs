using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Markup;

namespace XmlNamespaceReportDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the XML data file, the template document and the generated report.
            string xmlPath = Path.Combine(Environment.CurrentDirectory, "Persons.xml");
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // 1. Create an XML file that uses a namespace with a prefix.
            //    The root element declares the prefix "ns" that will be used for all child elements.
            string xmlContent =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<ns:persons xmlns:ns=""http://example.com/person"">
    <ns:person>
        <ns:name>John Doe</ns:name>
        <ns:age>30</ns:age>
    </ns:person>
    <ns:person>
        <ns:name>Jane Smith</ns:name>
        <ns:age>28</ns:age>
    </ns:person>
</ns:persons>";
            File.WriteAllText(xmlPath, xmlContent);

            // 2. Load the template document.
            //    The template contains a repeating region that iterates over "persons.person"
            //    and displays the "name" and "age" fields.
            Document template = new Document(templatePath);

            // 3. Create an XmlDataSource from the XML file.
            //    The constructor XmlDataSource(string) is provided by the library.
            XmlDataSource dataSource = new XmlDataSource(xmlPath);

            // 4. Build the report.
            //    Use the overload that accepts a data source name ("persons") so that the template
            //    can reference the XML root element directly.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "persons");

            // 5. (Optional) Demonstrate mapping a StructuredDocumentTag to a prefixed element.
            //    This shows how the prefix mapping string is used to resolve the XPath.
            //    The tag will display the name of the first person.
            CustomXmlPart xmlPart = template.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xmlContent);
            StructuredDocumentTag tag = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Block);
            // Prefix mapping declares the "ns" prefix used in the XPath expression.
            string prefixMapping = "xmlns:ns='http://example.com/person'";
            // XPath selects the first <ns:name> element.
            tag.XmlMapping.SetMapping(xmlPart, "/ns:persons[1]/ns:person[1]/ns:name[1]", prefixMapping);
            template.FirstSection.Body.AppendChild(tag);

            // 6. Save the generated report.
            template.Save(reportPath);
        }
    }
}
