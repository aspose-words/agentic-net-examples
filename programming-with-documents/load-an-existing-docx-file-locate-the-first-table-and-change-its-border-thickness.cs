using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the source and output documents.
        string sourcePath = "Sample.docx";
        string outputPath = "Modified.docx";

        // -----------------------------------------------------------------
        // Create a sample DOCX file with a simple table if it does not exist.
        // -----------------------------------------------------------------
        if (!File.Exists(sourcePath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            builder.Writeln("Sample document with a table:");
            builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Cell 1");
            builder.InsertCell();
            builder.Writeln("Cell 2");
            builder.EndRow();
            builder.EndTable();

            sampleDoc.Save(sourcePath);
        }

        // -------------------------------------------------
        // Load the existing document.
        // -------------------------------------------------
        Document doc = new Document(sourcePath);

        // -------------------------------------------------
        // Locate the first table in the document.
        // -------------------------------------------------
        Table firstTable = null;
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            firstTable = doc.FirstSection.Body.Tables[0];
        }

        // -------------------------------------------------
        // Change the border thickness of the table (2 points).
        // -------------------------------------------------
        if (firstTable != null)
        {
            // Apply a single line style with a thickness of 2 points and black color to all borders.
            firstTable.SetBorders(LineStyle.Single, 2.0, Color.Black);
        }

        // -------------------------------------------------
        // Save the modified document.
        // -------------------------------------------------
        doc.Save(outputPath);
    }
}
