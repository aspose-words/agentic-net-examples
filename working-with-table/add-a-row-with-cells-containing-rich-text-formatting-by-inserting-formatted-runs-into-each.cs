using System;
using System.IO;
using System.Drawing;
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

            // Start the table.
            builder.StartTable();

            // First row – simple header cells.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row – cells with rich formatted runs.

            // First cell.
            Cell cell1 = builder.InsertCell();
            Paragraph para1 = cell1.FirstParagraph;

            Run runBold = new Run(doc, "Bold");
            runBold.Font.Bold = true;
            para1.AppendChild(runBold);

            Run runNormal = new Run(doc, " Normal");
            para1.AppendChild(runNormal);

            // Second cell.
            Cell cell2 = builder.InsertCell();
            Paragraph para2 = cell2.FirstParagraph;

            Run runItalic = new Run(doc, "Italic");
            runItalic.Font.Italic = true;
            para2.AppendChild(runItalic);

            Run runRed = new Run(doc, " Red");
            // Correct property to set font color.
            runRed.Font.Color = Color.Red;
            para2.AppendChild(runRed);

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "FormattedTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");
        }
    }
}
