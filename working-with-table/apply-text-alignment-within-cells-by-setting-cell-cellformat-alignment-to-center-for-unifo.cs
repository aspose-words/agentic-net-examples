using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Add 3 rows and 3 columns with sample text.
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Center-align the text inside each cell.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            foreach (Row r in table.Rows)
            {
                foreach (Cell cell in r.Cells)
                {
                    foreach (Paragraph para in cell.Paragraphs)
                    {
                        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    }
                }
            }
        }

        // Save the document.
        string fileName = "AlignedTable.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception("The output document was not created.");
    }
}
