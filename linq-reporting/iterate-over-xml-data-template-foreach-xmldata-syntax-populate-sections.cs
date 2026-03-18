using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsXmlForeachDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains the <<foreach [in xmlData]>> tag.
            // Example template content:
            //   <<foreach [in xmlData]>>
            //   Name: <<[Name]>>
            //   Age: <<[Age]>>
            //   <</foreach>>
            string templatePath = @"C:\Docs\TemplateWithForeach.docx";

            // Path to the XML file that will be used as the data source.
            // Example XML:
            //   <People>
            //     <Person>
            //       <Name>John Doe</Name>
            //       <Age>30</Age>
            //     </Person>
            //     <Person>
            //       <Name>Jane Smith</Name>
            //       <Age>25</Age>
            //     </Person>
            //   </People>
            string xmlPath = @"C:\Docs\People.xml";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create an XmlDataSource from the XML file using the default loading options.
            XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

            // Build the report. The third argument is the name by which the template
            // refers to the data source (the identifier used inside <<foreach>>).
            // In the template we used "xmlData", so we pass the same name here.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, xmlDataSource, "xmlData");

            // Save the populated document.
            string outputPath = @"C:\Docs\ReportFromXml.docx";
            doc.Save(outputPath);
        }
    }
}
