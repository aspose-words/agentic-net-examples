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

        // Build a simple 3x3 table.
        Table table = builder.StartTable();

        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Ensure every paragraph inside the table keeps together on the same page.
        // This mimics the effect of keeping each row together.
        foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
        {
            foreach (Paragraph para in cell.Paragraphs)
            {
                para.ParagraphFormat.KeepTogether = true;
            }
        }

        // Define output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableKeepTogether.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
