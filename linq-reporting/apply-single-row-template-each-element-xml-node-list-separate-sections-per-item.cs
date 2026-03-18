using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1.  Define a repeating section that will be duplicated for each XML item.
        // -----------------------------------------------------------------
        // The repeating section is a StructuredDocumentTag (SDT) with type RepeatingSection
        // and markup level set to Block, so each iteration will produce a whole new section.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        // Append the repeating section to the body of the first (and only) section.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Inside the repeating section we need a RepeatingSectionItem that represents one iteration.
        StructuredDocumentTag item = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
        repeatingSection.AppendChild(item);

        // Add a paragraph with a placeholder that will be replaced by data from the XML.
        Paragraph para = new Paragraph(doc);
        item.AppendChild(para);
        // The placeholder syntax "<<[ds.Title]>>" tells the ReportingEngine to insert the value of the
        // <Title> element of the current XML node (the data source is named "ds").
        Run run = new Run(doc, "<<[ds.Title]>>");
        para.AppendChild(run);

        // -----------------------------------------------------------------
        // 2.  Prepare the XML data source.
        // -----------------------------------------------------------------
        // Example XML containing a list of <item> elements.
        string xml = @"
            <items>
                <item><Title>First Section Title</Title></item>
                <item><Title>Second Section Title</Title></item>
                <item><Title>Third Section Title</Title></item>
            </items>";

        // Create an XmlDataSource from a memory stream (no file I/O required).
        using (MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
        {
            XmlDataSource dataSource = new XmlDataSource(xmlStream);

            // -----------------------------------------------------------------
            // 3.  Build the report – the engine will duplicate the repeating section
            //     once for each <item> node and replace the placeholder with the
            //     corresponding <Title> value.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The third argument ("ds") is the name used in the template placeholders.
            engine.BuildReport(doc, dataSource, "ds");
        }

        // -----------------------------------------------------------------
        // 4.  Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save("Output.docx");
    }
}
