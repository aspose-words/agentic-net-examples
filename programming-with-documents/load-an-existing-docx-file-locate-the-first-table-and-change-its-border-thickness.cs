using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for input and output files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Paths for the sample source document and the result document.
        string sourcePath = Path.Combine(dataDir, "Source.docx");
        string resultPath = Path.Combine(dataDir, "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX file that contains a simple table.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Build a 2x2 table.
        Table table = builder.StartTable();
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
        sampleDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the existing document, locate the first table, and modify its borders.
        // -----------------------------------------------------------------
        Document doc = new Document(sourcePath);

        // Ensure the document contains at least one table.
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            Table firstTable = doc.FirstSection.Body.Tables[0];

            // Change all borders to a single blue line with a thickness of 2 points.
            firstTable.SetBorders(LineStyle.Single, 2.0, Color.Blue);
        }

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(resultPath);
    }
}
