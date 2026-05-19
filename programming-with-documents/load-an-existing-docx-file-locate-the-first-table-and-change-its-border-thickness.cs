using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define file names for the sample and the result.
        const string samplePath = "Sample.docx";
        const string resultPath = "Modified.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains a simple table.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Start a table with two rows and two columns.
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

        // Save the sample document to the file system.
        sampleDoc.Save(samplePath);

        // -----------------------------------------------------------------
        // 2. Load the existing document from disk.
        // -----------------------------------------------------------------
        Document doc = new Document(samplePath);

        // -----------------------------------------------------------------
        // 3. Locate the first table in the document.
        // -----------------------------------------------------------------
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // -----------------------------------------------------------------
        // 4. Change the border thickness of the table.
        //    Here we set all borders to a single line with a width of 2 points.
        // -----------------------------------------------------------------
        firstTable.SetBorders(LineStyle.Single, 2.0, Color.Black);

        // -----------------------------------------------------------------
        // 5. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(resultPath);
    }
}
