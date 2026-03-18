using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a sample table to demonstrate the alignment.
        Table sampleTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Retrieve the left and right page margins from the first section.
        double leftPageMargin = doc.FirstSection.PageSetup.LeftMargin;
        double rightPageMargin = doc.FirstSection.PageSetup.RightMargin;

        // Align every table in the document to the page margins.
        foreach (Table tbl in doc.GetChildNodes(NodeType.Table, true))
        {
            // Position the table's left edge at the left page margin.
            tbl.LeftIndent = leftPageMargin;

            // Ensure the table's right edge does not exceed the right page margin.
            // DistanceRight defines the space between the table's right border and surrounding text.
            tbl.DistanceRight = rightPageMargin;
        }

        // Save the modified document.
        doc.Save("AlignedTables.docx");
    }
}
