using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the input and output documents.
        const string inputPath = "SampleTable.docx";
        const string outputPath = "ModifiedTable.docx";

        // -----------------------------------------------------------------
        // Create a sample DOCX file containing a simple table.
        // This ensures the example works even when no external file exists.
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
        sampleDoc.Save(inputPath);

        // ---------------------------------------------------------------
        // Load the existing document, locate the first table, and modify it.
        // ---------------------------------------------------------------
        Document doc = new Document(inputPath);

        // Retrieve the first table in the document.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Set all borders of the table to a single line, 3 points thick, black color.
        firstTable.SetBorders(LineStyle.Single, 3.0, Color.Black);

        // Save the modified document.
        doc.Save(outputPath);
    }
}
