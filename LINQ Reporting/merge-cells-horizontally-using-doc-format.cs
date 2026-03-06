using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExamples
{
    class HorizontalCellMerge
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder which simplifies construction of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            builder.StartTable();

            // First cell – mark it as the first cell in a horizontally merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("This text spans two merged cells.");

            // Second cell – merge it with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the first row.
            builder.EndRow();

            // Add a normal (unmerged) cell in the second row to demonstrate mixed content.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None; // Reset merge setting.
            builder.Write("Unmerged cell in second row.");

            // End the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document in DOC format.
            doc.Save("CellFormat.HorizontalMerge.doc");
        }
    }
}
