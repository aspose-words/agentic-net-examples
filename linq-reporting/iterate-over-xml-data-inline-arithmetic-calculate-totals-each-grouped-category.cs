using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsXmlReport
{
    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains the reporting tags.
            // The template should have inline arithmetic expressions, e.g. <<[Category.Total]>> or
            // <<[Amount] + [Tax]>> inside a <<foreach>> block to calculate totals per group.
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Path to the XML file that holds the source data.
            // Example XML structure:
            // <Report>
            //   <Category name="Food">
            //     <Item amount="10.5" tax="0.5" />
            //     <Item amount="20.0" tax="1.0" />
            //   </Category>
            //   <Category name="Books">
            //     <Item amount="15.0" tax="0.75" />
            //   </Category>
            // </Report>
            string xmlDataPath = @"C:\Data\ReportData.xml";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create an XML data source from the file.
            // This uses the provided XmlDataSource constructor (string).
            XmlDataSource xmlDataSource = new XmlDataSource(xmlDataPath);

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report.
            // The third parameter is the name by which the XML data source will be referenced in the template.
            // In the template you can use tags like <<foreach [in Report.Category]>> and
            // inline arithmetic such as <<[Item.amount] + [Item.tax]>> to compute totals.
            engine.BuildReport(doc, xmlDataSource, "Report");

            // Save the generated report.
            // This uses the provided Document.Save(string) method.
            string outputPath = @"C:\Output\GeneratedReport.docx";
            doc.Save(outputPath);
        }
    }
}
