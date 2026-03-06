using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document – this will be used to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table. The first call to InsertCell will also create the first row.
        builder.StartTable();

        // Insert a new cell into the current row and obtain the Cell object.
        Cell cell = builder.InsertCell();

        // The builder's cursor is now inside the newly created cell.
        // Write the desired text into the cell.
        builder.Write("Text inserted into the cell.");

        // Finish the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document as a Markdown file.
        doc.Save("CellOutput.md", SaveFormat.Markdown);
    }
}
