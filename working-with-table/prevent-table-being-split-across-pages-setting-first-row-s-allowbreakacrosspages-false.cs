using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableRowBreakControl
{
    class Program
    {
        static void Main()
        {
            // Load an existing document that contains at least one table.
            Document doc = new Document("Input.docx");

            // Get the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // Access the first row of the table.
            Row firstRow = table.FirstRow;

            // Prevent this row from being split across pages.
            // Setting AllowBreakAcrossPages to false keeps the row together.
            firstRow.RowFormat.AllowBreakAcrossPages = false;

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
