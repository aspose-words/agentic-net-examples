using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define directories and file names.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        string sourcePath = Path.Combine(dataDir, "Source.docx");
        string outputPath = Path.Combine(dataDir, "Modified.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains a simple table.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

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

        // Save the sample document to disk.
        sampleDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the existing document, locate the first table, and modify its borders.
        // -----------------------------------------------------------------
        Document doc = new Document(sourcePath);

        // Retrieve the first table in the document.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Change all borders of the table to a single black line with a thickness of 3 points.
        firstTable.SetBorders(LineStyle.Single, 3.0, Color.Black);

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);

        // Optional: confirm that the output file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Modified document saved to: {outputPath}");
        }
    }
}
