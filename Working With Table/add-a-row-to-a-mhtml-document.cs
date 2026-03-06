using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    class AddRowToMhtml
    {
        static void Main()
        {
            // Load the existing MHTML document.
            Document doc = new Document("input.mhtml");

            // Ensure the document contains at least one table.
            if (doc.FirstSection?.Body?.Tables?.Count > 0)
            {
                // Get the first table in the document.
                Table table = doc.FirstSection.Body.Tables[0];

                // Create a new row that belongs to the same document.
                Row newRow = new Row(doc);

                // Ensure the row has at least one cell.
                newRow.EnsureMinimum();

                // Add content to the first cell of the new row.
                Cell firstCell = newRow.FirstCell;
                Paragraph para = new Paragraph(doc);
                firstCell.FirstParagraph?.RemoveAllChildren(); // Clear any existing empty paragraph.
                firstCell.AppendChild(para);
                para.AppendChild(new Run(doc, "New row added via code."));

                // Append the new row to the table.
                table.AppendChild(newRow);
            }
            else
            {
                Console.WriteLine("No tables found in the document.");
            }

            // Save the modified document back to MHTML format.
            doc.Save("output.mhtml");
        }
    }
}
