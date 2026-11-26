using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Paths for the sample and output documents.
        string sourcePath = "Sample.docx";
        string outputPath = "Modified.docx";

        // Ensure any previous files are removed.
        if (File.Exists(sourcePath)) File.Delete(sourcePath);
        if (File.Exists(outputPath)) File.Delete(outputPath);

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX file containing a simple table.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with 2 rows and 2 columns.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Save the sample document.
        doc.Save(sourcePath);

        // -----------------------------------------------------------------
        // Step 2: Load the existing DOCX file.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // Step 3: Locate the first table in the document.
        // -----------------------------------------------------------------
        Table firstTable = loadedDoc.GetChild(NodeType.Table, 0, true) as Table;
        if (firstTable != null)
        {
            // -----------------------------------------------------------------
            // Step 4: Change the border thickness of the table.
            // -----------------------------------------------------------------
            // Set all borders to a single line with a thickness of 2 points.
            firstTable.SetBorders(LineStyle.Single, 2.0, System.Drawing.Color.Black);
        }

        // -----------------------------------------------------------------
        // Step 5: Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath);

        // Optional verification that the output file exists.
        if (File.Exists(outputPath))
        {
            // The program finishes here without any user interaction.
        }
    }
}
