using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a source document and define a custom table style.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        Style customStyle = sourceDoc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 10;
        // Note: Shading is not set because the Style class does not expose a Shading property in this version.

        // Build a simple table in the source document and apply the custom style.
        sourceBuilder.StartTable();
        sourceBuilder.InsertCell();
        sourceBuilder.Writeln("Header 1");
        sourceBuilder.InsertCell();
        sourceBuilder.Writeln("Header 2");
        sourceBuilder.EndRow();

        sourceBuilder.InsertCell();
        sourceBuilder.Writeln("Cell 1");
        sourceBuilder.InsertCell();
        sourceBuilder.Writeln("Cell 2");
        sourceBuilder.EndRow();
        sourceBuilder.EndTable();

        Table sourceTable = (Table)sourceDoc.GetChildNodes(NodeType.Table, true)[0];
        sourceTable.Style = customStyle;

        // Save the source document to a memory stream (optional step to simulate file I/O).
        using (MemoryStream ms = new MemoryStream())
        {
            sourceDoc.Save(ms, SaveFormat.Docx);
            ms.Position = 0;

            // Load a destination document (empty) and copy the custom table style from the source.
            Document destDoc = new Document();
            Style sourceStyle = sourceDoc.Styles["MyCustomTableStyle"];
            destDoc.Styles.AddCopy(sourceStyle);

            // Build a table in the destination document and apply the copied style.
            DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
            destBuilder.StartTable();
            destBuilder.InsertCell();
            destBuilder.Writeln("Dest Header 1");
            destBuilder.InsertCell();
            destBuilder.Writeln("Dest Header 2");
            destBuilder.EndRow();

            destBuilder.InsertCell();
            destBuilder.Writeln("Dest Cell 1");
            destBuilder.InsertCell();
            destBuilder.Writeln("Dest Cell 2");
            destBuilder.EndRow();
            destBuilder.EndTable();

            Table destTable = (Table)destDoc.GetChildNodes(NodeType.Table, true)[0];
            destTable.Style = destDoc.Styles["MyCustomTableStyle"];

            // Save the destination document.
            string outputPath = "Destination.docx";
            destDoc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The destination document was not created.");
        }
    }
}
