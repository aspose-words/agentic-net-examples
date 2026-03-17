using System;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

class RemoveTableBorders
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Get the first table in the document (adjust index as needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Remove all borders from the table.
        // This sets each border's LineStyle to LineStyle.None internally.
        table.ClearBorders();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
