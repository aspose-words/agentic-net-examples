using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define file paths relative to the executable directory.
        string samplePath = "Sample.docx";
        string outputPath = "Modified.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX file containing a simple 2x2 table.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Start a table and add two rows with two cells each.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Save the sample document.
        sampleDoc.Save(samplePath);

        // -----------------------------------------------------------------
        // 2. Load the existing document, locate the first table, and modify its borders.
        // -----------------------------------------------------------------
        Document doc = new Document(samplePath);

        // Ensure the document contains at least one table.
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            // Get the first table in the document.
            Table firstTable = doc.FirstSection.Body.Tables[0];

            // Change all borders to a single line with a thickness of 2 points and black color.
            firstTable.SetBorders(LineStyle.Single, 2.0, Color.Black);
        }

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
