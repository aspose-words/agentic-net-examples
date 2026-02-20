using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableAlignmentExample
{
    static void Main()
    {
        // Load an existing WORDML document.
        Document doc = new Document("InputDocument.xml");

        // Get the first table in the document (adjust index as needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Center the table on the page.
        table.Alignment = TableAlignment.Center;

        // Set the table width to 100% of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(100);

        // Save the modified document back to WORDML format.
        doc.Save("OutputDocument.xml");
    }
}
