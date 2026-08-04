using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for the documents.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Paths for the source and the modified documents.
        string sourcePath = Path.Combine(dataDir, "Source.docx");
        string outputPath = Path.Combine(dataDir, "Modified.docx");

        // -----------------------------------------------------------------
        // Create a sample DOCX file that contains a simple table.
        // -----------------------------------------------------------------
        Document createDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createDoc);

        // Build a 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Save the sample document.
        createDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // Load the existing document, locate the first table, and change its border thickness.
        // -----------------------------------------------------------------
        Document loadDoc = new Document(sourcePath);

        // Ensure the document contains at least one table.
        if (loadDoc.FirstSection?.Body?.Tables?.Count > 0)
        {
            Table firstTable = loadDoc.FirstSection.Body.Tables[0];

            // Set all borders of the table to a single line with a thickness of 3 points.
            firstTable.SetBorders(LineStyle.Single, 3.0, Color.Black);
        }

        // Save the modified document.
        loadDoc.Save(outputPath);
    }
}
