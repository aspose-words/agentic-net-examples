using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

class EnableInvariantXmlNumbers
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple MERGEFIELD that will display a numeric value.
        builder.InsertField("MERGEFIELD Amount \\# $#,##0.00");

        // Change the current thread culture to a culture that uses a comma as decimal separator.
        // This will demonstrate that without invariant formatting the number would be parsed incorrectly.
        Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

        // Enable invariant culture number format for all fields in the document.
        // This makes the reporting engine interpret numeric strings using the invariant culture ('.' as decimal separator).
        doc.FieldOptions.UseInvariantCultureNumberFormat = true;

        // Prepare XML data as a string. Numeric values are formatted using invariant culture (dot as decimal separator).
        string xmlData = @"
            <Root>
                <Record>
                    <Amount>1234.56</Amount>
                </Record>
            </Root>";

        // Load the XML data into a stream.
        using (MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlData)))
        {
            // Create an XmlDataSource from the stream. No schema is required.
            XmlDataSource dataSource = new XmlDataSource(xmlStream);

            // Build the report using the XML data source.
            // ReportingEngine.BuildReport is an instance method, so we need to create an engine object.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "Root");

            // Save the resulting document.
            doc.Save("Result.docx");
        }
    }
}
