using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

class InTableListFromXml
{
    static void Main()
    {
        // Paths to the XML data file and the output document.
        const string xmlPath = "Data.xml";
        const string outputPath = "Result.docx";

        // Load the XML document.
        XDocument xDoc = XDocument.Load(xmlPath);

        // Assume the XML structure is:
        // <Root>
        //   <Item>
        //     <Field1>...</Field1>
        //     <Field2>...</Field2>
        //   </Item>
        //   ...
        // </Root>
        // The first <Item> element defines the column names.
        XElement firstItem = xDoc.Root.Elements().FirstOrDefault()
                           ?? throw new InvalidOperationException("No data elements found.");

        // Extract the element names of the first item – these become the table headers.
        string[] columnNames = firstItem.Elements()
                                        .Select(e => e.Name.LocalName)
                                        .ToArray();

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ---------- Header row ----------
        foreach (string col in columnNames)
        {
            builder.InsertCell();
            builder.Write(col);               // Write the column name as plain text.
        }
        builder.EndRow();                     // Finish the header row.

        // ---------- Data row ----------
        // Insert placeholders that the ReportingEngine will replace with XML values.
        foreach (string col in columnNames)
        {
            builder.InsertCell();
            // ReportingEngine uses the <<[FieldName]>> syntax.
            builder.Write($"<<[{col}]>>");
        }
        builder.EndRow();                     // Finish the data row.

        // End the table.
        builder.EndTable();

        // Build the report using the XML data source.
        // The third argument is the name used in the template to reference the data source.
        // Here we use the root element name (e.g., "Root").
        XmlDataSource dataSource = new XmlDataSource(xmlPath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, xDoc.Root.Name.LocalName);

        // Save the populated document.
        doc.Save(outputPath);
    }
}
