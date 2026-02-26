using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoCell
{
    static void Main()
    {
        // Path to the folder that contains the template DOT file.
        string dataDir = @"C:\Data\";
        string templatePath = dataDir + "Template.dot";

        // Load the DOT template document.
        Document doc = new Document(templatePath);

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the first cell of the first table.
        // Parameters: tableIndex, rowIndex, columnIndex, cellIndex.
        // Here we target the cell at row 0, column 0 of the first table.
        builder.MoveToCell(0, 0, 0, 0);

        // Insert the desired text into the cell.
        builder.Write("Inserted text into the cell.");

        // Save the modified document. You can save as DOCX or any other supported format.
        string outputPath = dataDir + "Result.docx";
        doc.Save(outputPath);
    }
}
