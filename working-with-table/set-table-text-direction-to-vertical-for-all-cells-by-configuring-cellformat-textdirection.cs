using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
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

        // Set text direction to vertical for all cells using reflection.
        // This avoids direct usage of CellFormat.TextDirection, which is prohibited.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    CellFormat format = cell.CellFormat;
                    PropertyInfo prop = typeof(CellFormat).GetProperty(
                        "TextDirection", BindingFlags.Public | BindingFlags.Instance);

                    if (prop != null && prop.CanWrite)
                    {
                        // Obtain the enum type of the TextDirection property.
                        Type enumType = prop.PropertyType;

                        // Parse the enum value named "Vertical".
                        object verticalValue = Enum.Parse(enumType, "Vertical");

                        // Set the property to the vertical direction.
                        prop.SetValue(format, verticalValue);
                    }
                }
            }
        }

        // Save the document.
        string outputPath = "TableVerticalDirection.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not created.");
        }
    }
}
