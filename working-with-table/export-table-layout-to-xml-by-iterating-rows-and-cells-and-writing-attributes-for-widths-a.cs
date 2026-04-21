using System;
using System.IO;
using System.Xml.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.CellFormat.Width = 100;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder.Write("Cell 1,1");

        // First row, second cell.
        builder.InsertCell();
        builder.CellFormat.Width = 150;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.CellFormat.Width = 120;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
        builder.Write("Cell 2,1");

        // Second row, second cell.
        builder.InsertCell();
        builder.CellFormat.Width = 130;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document (optional, for verification).
        string docPath = "SampleTable.docx";
        doc.Save(docPath);

        // Export table layout to XML.
        XDocument xml = new XDocument(new XElement("Tables"));
        NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);
        int tableIndex = 0;

        foreach (Table tbl in tableNodes)
        {
            XElement tblElement = new XElement("Table", new XAttribute("Index", tableIndex));
            int rowIndex = 0;

            foreach (Row row in tbl.Rows)
            {
                XElement rowElement = new XElement("Row", new XAttribute("Index", rowIndex));
                int cellIndex = 0;

                foreach (Cell cell in row.Cells)
                {
                    // Cell width (points).
                    double width = cell.CellFormat.Width;

                    // Background color as HTML hex string.
                    Color bgColor = cell.CellFormat.Shading.BackgroundPatternColor;
                    string bgHex = bgColor.IsEmpty ? "" : ColorTranslator.ToHtml(bgColor);

                    // Cell text.
                    string text = cell.ToString(SaveFormat.Text).Trim();

                    XElement cellElement = new XElement("Cell",
                        new XAttribute("Index", cellIndex),
                        new XAttribute("Width", width),
                        new XAttribute("BackgroundColor", bgHex),
                        new XAttribute("Text", text));

                    rowElement.Add(cellElement);
                    cellIndex++;
                }

                tblElement.Add(rowElement);
                rowIndex++;
            }

            xml.Root.Add(tblElement);
            tableIndex++;
        }

        // Save the XML report.
        string xmlPath = "TableLayout.xml";
        xml.Save(xmlPath);

        // Verify that the XML file was created.
        if (!File.Exists(xmlPath))
            throw new Exception("Failed to create the XML report.");

        // Inform the user where the report was saved.
        Console.WriteLine($"XML report saved to {Path.GetFullPath(xmlPath)}");
    }
}
