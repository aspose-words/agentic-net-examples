using System;
using System.IO;
using System.Drawing;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table with some formatting.
        Table table = builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

        builder.InsertCell();
        builder.Write("Header 2");
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        builder.EndRow();

        // Second row – data cells.
        builder.InsertCell();
        builder.Write("Cell A");
        builder.CellFormat.Shading.BackgroundPatternColor = Color.White;

        builder.InsertCell();
        builder.Write("Cell B");
        builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
        builder.EndRow();

        builder.EndTable();

        // Save the document to disk.
        string docPath = "SampleTable.docx";
        doc.Save(docPath);

        // Verify that the document was saved.
        if (!File.Exists(docPath))
            throw new InvalidOperationException("Failed to save the Word document.");

        // Export the layout of all tables to an XML file.
        XDocument xmlDoc = new XDocument(new XElement("Tables"));
        int tableIndex = 0;

        foreach (Table tbl in doc.GetChildNodes(NodeType.Table, true))
        {
            XElement tblElement = new XElement("Table",
                new XAttribute("Index", tableIndex),
                new XAttribute("Title", tbl.Title ?? string.Empty),
                new XAttribute("Description", tbl.Description ?? string.Empty));

            int rowIndex = 0;
            foreach (Row row in tbl.Rows)
            {
                XElement rowElement = new XElement("Row",
                    new XAttribute("Index", rowIndex));

                int cellIndex = 0;
                foreach (Cell cell in row.Cells)
                {
                    // Retrieve cell width (if set) and background color.
                    double width = cell.CellFormat.Width;
                    Color bgColor = cell.CellFormat.Shading.BackgroundPatternColor;

                    XElement cellElement = new XElement("Cell",
                        new XAttribute("Index", cellIndex),
                        new XAttribute("Width", width),
                        new XAttribute("BackgroundColor", ColorToHex(bgColor)));

                    // Add the cell's text content.
                    string cellText = cell.ToString(SaveFormat.Text).Trim();
                    cellElement.Add(new XElement("Text", cellText));

                    rowElement.Add(cellElement);
                    cellIndex++;
                }

                tblElement.Add(rowElement);
                rowIndex++;
            }

            xmlDoc.Root.Add(tblElement);
            tableIndex++;
        }

        // Save the XML representation.
        string xmlPath = "TableLayout.xml";
        xmlDoc.Save(xmlPath);

        // Verify that the XML file was created.
        if (!File.Exists(xmlPath))
            throw new InvalidOperationException("Failed to export table layout to XML.");
    }

    // Helper method to convert a Color to a hex string (e.g., #RRGGBB).
    private static string ColorToHex(Color color)
    {
        if (color.IsEmpty)
            return string.Empty;
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
    }
}
