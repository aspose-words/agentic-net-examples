using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsFieldUpdateExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing document that contains a table with formula fields.
            // (Replace the path with the actual location of your document.)
            Document doc = new Document(@"C:\Docs\InputWithTable.docx");

            // Find the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // Insert a new row at the end of the table.
            // The new row will inherit the formatting of the previous row.
            Row newRow = (Row)table.LastRow.Clone(true);
            table.Rows.Add(newRow);

            // Optionally, fill the new row with data.
            // For example, set the first cell's text to a numeric value that a formula field will use.
            newRow.Cells[0].FirstParagraph.AppendChild(new Run(doc, "42"));
            // If the table contains formula fields (e.g., = SUM(ABOVE)), they will reference the new data.

            // Recalculate all fields in the document, including the formula fields in the table.
            doc.UpdateFields();

            // Save the updated document.
            doc.Save(@"C:\Docs\OutputWithUpdatedFields.docx");
        }
    }
}
