using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing; // Required for Shading class

class Program
{
    static void Main()
    {
        // Load the existing document.
        Document doc = new Document("Input.docx");

        // Iterate through all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Iterate through each row of the current table.
            foreach (Row row in table.Rows)
            {
                bool shouldHighlight = false;

                // Check each cell in the row for the <<setBackground>> tag.
                foreach (Cell cell in row.Cells)
                {
                    if (cell.GetText().Contains("<<setBackground>>"))
                    {
                        shouldHighlight = true;
                        break;
                    }
                }

                // If the tag was found, set the background of every cell in the row to yellow.
                if (shouldHighlight)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        // CellFormat.Shading controls the background shading of a cell.
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
                    }
                }
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
