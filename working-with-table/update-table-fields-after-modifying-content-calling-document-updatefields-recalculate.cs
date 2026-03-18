using System;
using Aspose.Words;
using Aspose.Words.Tables;

class UpdateTableFieldsExample
{
    static void Main()
    {
        // Load an existing document that contains a table with fields.
        Document doc = new Document("Input.docx");

        // Find the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Modify the content of a specific cell (row 0, column 0) – for example, change a placeholder text.
        Cell cell = table.Rows[0].Cells[0];
        cell.FirstParagraph.AppendChild(new Run(doc, "Updated value"));

        // After modifying the document, update all fields so that any calculations or references
        // inside the table (e.g., =, NUMCHARS, NUMWORDS, etc.) are recalculated.
        doc.UpdateFields();

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
