using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the Word template that contains the reporting tags.
        // The template should include a groupby expression, e.g.:
        // <<foreach [groupby items/category]>>Category: <<[Key]>> Count: <<[Count]>> <</foreach>>
        Document doc = new Document("Template.docx");

        // Load the XML data source. The XML file should have a repeating element named "item"
        // with a child element or attribute called "category".
        XmlDataSource xmlData = new XmlDataSource("Data.xml");

        // Build the report by merging the XML data into the template.
        // The third parameter is the name used in the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, xmlData, "items");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
