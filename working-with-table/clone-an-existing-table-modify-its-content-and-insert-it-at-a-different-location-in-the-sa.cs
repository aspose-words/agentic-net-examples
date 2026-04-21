using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an introductory paragraph.
        builder.Writeln("Document start.");

        // Build the original table.
        Table originalTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("Original Cell 1");
        builder.InsertCell();
        builder.Write("Original Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Original Cell 3");
        builder.InsertCell();
        builder.Write("Original Cell 4");
        builder.EndRow();
        builder.EndTable(); // Ends the table and leaves the cursor after it.

        // Add a paragraph that will separate the two tables.
        builder.Writeln("Paragraph between tables.");
        // Capture the paragraph node so we know where to insert the cloned table.
        Paragraph separatorParagraph = builder.CurrentParagraph;

        // Clone the original table (deep clone).
        Table clonedTable = (Table)originalTable.Clone(true);

        // Modify the content of the cloned table.
        // Change the text of the first cell.
        Run firstRun = clonedTable.FirstRow.FirstCell.FirstParagraph.Runs[0];
        firstRun.Text = "Cloned Cell 1";

        // Insert the cloned table after the separator paragraph.
        separatorParagraph.ParentNode.InsertAfter(clonedTable, separatorParagraph);

        // Save the resulting document.
        string outputPath = "ClonedTable.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
