using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source document that already contains the table style we want to copy.
        Document srcDoc = new Document("Source.docx");

        // Retrieve the table style by its name. Replace "MyTableStyle1" with the actual style name.
        Style srcStyle = srcDoc.Styles["MyTableStyle1"];
        if (srcStyle == null)
        {
            Console.WriteLine("The specified table style was not found in the source document.");
            return;
        }

        // Create a new (empty) destination document.
        Document dstDoc = new Document();

        // Copy the table style from the source document into the destination document.
        // The AddCopy method creates a copy and automatically generates a unique name if needed.
        Style copiedStyle = dstDoc.Styles.AddCopy(srcStyle);

        // Optionally rename the copied style for clarity.
        copiedStyle.Name = "MyTableStyle1_Copy";

        // Build a simple 2‑cell table in the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Apply the copied table style to the newly created table.
        table.Style = copiedStyle;

        // Save the destination document with the imported style applied.
        dstDoc.Save("Destination.docx");
    }
}
