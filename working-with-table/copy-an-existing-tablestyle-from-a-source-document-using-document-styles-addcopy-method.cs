using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the source and destination documents.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string destinationPath = Path.Combine(outputDir, "Destination.docx");

        // ---------- Create source document with a custom table style ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // Define a new table style.
        string styleName = "MyCustomTableStyle";
        TableStyle tableStyle = (TableStyle)sourceDoc.Styles.Add(StyleType.Table, styleName);
        // Set some simple formatting for the style.
        tableStyle.Shading.BackgroundPatternColor = System.Drawing.Color.LightBlue;
        tableStyle.Borders.Color = System.Drawing.Color.DarkBlue;
        tableStyle.Borders.LineWidth = 1.5;

        // Build a table and apply the custom style.
        srcBuilder.StartTable();
        srcBuilder.RowFormat.Height = 20;
        srcBuilder.InsertCell();
        srcBuilder.Write("Header 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Header 2");
        srcBuilder.EndRow();

        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        // Apply the style to the created table.
        Table srcTable = (Table)sourceDoc.GetChild(NodeType.Table, 0, true);
        srcTable.Style = tableStyle;

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // ---------- Load destination document and copy the table style ----------
        Document destDoc = new Document(); // start with an empty document
        // Copy the custom table style from source to destination.
        Style copiedStyle = destDoc.Styles.AddCopy(sourceDoc.Styles[styleName]);

        // Build a new table in the destination document using the copied style.
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.StartTable();
        destBuilder.InsertCell();
        destBuilder.Write("Dest Header 1");
        destBuilder.InsertCell();
        destBuilder.Write("Dest Header 2");
        destBuilder.EndRow();

        destBuilder.InsertCell();
        destBuilder.Write("Dest Cell 1");
        destBuilder.InsertCell();
        destBuilder.Write("Dest Cell 2");
        destBuilder.EndRow();
        destBuilder.EndTable();

        // Apply the copied style to the new table.
        Table destTable = (Table)destDoc.GetChild(NodeType.Table, 0, true);
        destTable.Style = copiedStyle;

        // Save the destination document.
        destDoc.Save(destinationPath);

        // Validate that the destination file was created.
        if (!File.Exists(destinationPath))
            throw new Exception("Destination document was not saved correctly.");

        // Optionally, clean up source file if not needed.
        // File.Delete(sourcePath);
    }
}
