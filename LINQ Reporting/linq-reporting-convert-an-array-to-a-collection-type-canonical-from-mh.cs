using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class LinqReportingArrayToCollection
{
    static void Main()
    {
        // Load a document that is stored in MHTML format.
        // The Document constructor handles loading from a file path.
        Document srcDoc = new Document("Template.mht");

        // Retrieve all paragraph nodes from the source document.
        // GetChildNodes returns a live NodeCollection.
        NodeCollection paragraphs = srcDoc.GetChildNodes(NodeType.Paragraph, true);

        // Convert the live collection to a fixed‑size array.
        // This uses the built‑in ToArray method of NodeCollection.
        Node[] paragraphArray = paragraphs.ToArray();

        // The LINQ Reporting Engine expects a data source that can be enumerated.
        // An array implements IEnumerable, so it can be used directly.
        // For demonstration we pass the array as the sole data source.
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the original template document and the array data source.
        // The overload with (Document, object) is used.
        engine.BuildReport(srcDoc, paragraphArray);

        // Save the populated document in DOCX format.
        srcDoc.Save("ReportFromMhtml.docx", SaveFormat.Docx);
    }
}
