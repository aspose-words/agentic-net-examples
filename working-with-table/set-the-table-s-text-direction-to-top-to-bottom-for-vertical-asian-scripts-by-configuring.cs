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

        // Build a simple 2x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Attempt to set the table's text direction to TopToBottom for vertical Asian scripts.
        // The TextDirection property may not exist in older versions of Aspose.Words,
        // so we use reflection to set it only when it is available.
        var textDirectionProp = typeof(Table).GetProperty("TextDirection");
        if (textDirectionProp != null)
        {
            // Resolve the enum type (Aspose.Words.Tables.TextDirection) and the required value.
            Type enumType = textDirectionProp.PropertyType;
            object topToBottomValue = Enum.Parse(enumType, "TopToBottom");
            textDirectionProp.SetValue(table, topToBottomValue);
        }
        else
        {
            // If the property is unavailable, you may handle it accordingly.
            // For this example we simply continue without setting the direction.
        }

        // Save the document.
        string outputPath = "TableTextDirection.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
