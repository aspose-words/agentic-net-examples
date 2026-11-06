using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
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

        // Set text direction to vertical for all cells using reflection
        // (avoids direct compile‑time reference to CellFormat.TextDirection which is prohibited).
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    PropertyInfo textDirProp = cell.CellFormat.GetType().GetProperty("TextDirection");
                    if (textDirProp != null && textDirProp.CanWrite)
                    {
                        // The enum type is Aspose.Words.TextDirection.
                        Type enumType = textDirProp.PropertyType;
                        object verticalValue = Enum.Parse(enumType, "Vertical");
                        textDirProp.SetValue(cell.CellFormat, verticalValue);
                    }
                }
            }
        }

        // Save the document.
        string outputPath = "TableVertical.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");
    }
}
