using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing WORDML (or DOCX) document.
        Document doc = new Document("Input.docx");

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Align the table to the center of the page.
        table.Alignment = TableAlignment.Center;

        // Set the table width to 100 % of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(100);

        // Save the modified document back to WORDML format.
        doc.Save("Output.xml", SaveFormat.WordML);
    }
}
