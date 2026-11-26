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

        // Build a simple 2‑cell table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Attempt to set the table's text direction to TopToBottom if the property exists.
        PropertyInfo textDirProp = typeof(Table).GetProperty("TextDirection", BindingFlags.Public | BindingFlags.Instance);
        if (textDirProp != null && textDirProp.CanWrite)
        {
            Type enumType = textDirProp.PropertyType;
            // Use Enum.TryParse to avoid exceptions if the enum value name differs.
            if (Enum.TryParse(enumType, "TopToBottom", out object topToBottomValue))
            {
                textDirProp.SetValue(table, topToBottomValue);
            }
        }
        // If the property does not exist, we simply continue without setting it.

        // Save the document.
        string outputPath = "TableTextDirection.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("Failed to create the output document.", outputPath);
        }
    }
}
