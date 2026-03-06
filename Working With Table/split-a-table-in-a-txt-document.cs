using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class SplitTableToTxt
{
    static void Main()
    {
        // Load the source document that contains a table.
        Document srcDoc = new Document("SourceWithTable.docx");

        // Get the first table in the document.
        Table srcTable = srcDoc.FirstSection.Body.Tables[0];

        // Iterate through each row of the source table.
        for (int rowIndex = 0; rowIndex < srcTable.Rows.Count; rowIndex++)
        {
            // Create a new empty document for the current row.
            Document rowDoc = new Document();
            // Ensure the document has at least one section and one paragraph.
            rowDoc.EnsureMinimum();

            // Create a new table and add it to the document body.
            Table newTable = new Table(rowDoc);
            rowDoc.FirstSection.Body.AppendChild(newTable);

            // Clone the current row (including its cells and contents) and add it to the new table.
            Row clonedRow = (Row)srcTable.Rows[rowIndex].Clone(true);
            newTable.AppendChild(clonedRow);

            // Optional: adjust table formatting if needed.
            newTable.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Prepare TXT save options – preserve the table layout for readability.
            TxtSaveOptions txtOptions = new TxtSaveOptions
            {
                PreserveTableLayout = true,
                // Force page breaks to be kept as '\f' characters (optional).
                ForcePageBreaks = false
            };

            // Build the output file name, e.g., Row_1.txt, Row_2.txt, etc.
            string outFileName = $"TableRow_{rowIndex + 1}.txt";

            // Save the document containing only the single row as a plain‑text file.
            rowDoc.Save(outFileName, txtOptions);
        }
    }
}
