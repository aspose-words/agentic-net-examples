using System;
using System.IO;
using System.Drawing;               // Needed for Color
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table with two columns.
            Table table = builder.StartTable();

            // ----- Header row (plain text) -----
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // ----- New row with rich formatted text in each cell -----
            // First cell
            Cell cell1 = builder.InsertCell();
            // The first paragraph already exists in the cell.
            Paragraph para1 = cell1.FirstParagraph;

            // Run 1: bold text
            Run runBold = new Run(doc, "Bold");
            runBold.Font.Bold = true;
            para1.AppendChild(runBold);

            // Run 2: normal text
            Run runNormal = new Run(doc, " and normal");
            para1.AppendChild(runNormal);

            // Second cell
            Cell cell2 = builder.InsertCell();
            Paragraph para2 = cell2.FirstParagraph;

            // Run 1: italic, red text
            Run runItalicRed = new Run(doc, "Italic Red");
            runItalicRed.Font.Italic = true;
            // Use Font.Color instead of the non‑existent FontColor property.
            runItalicRed.Font.Color = Color.Red;
            para2.AppendChild(runItalicRed);

            // Run 2: normal text
            Run runPlain = new Run(doc, " text");
            para2.AppendChild(runPlain);

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // Inform that the process completed.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
