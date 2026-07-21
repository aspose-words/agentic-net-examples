using System;
using System.IO;
using System.Drawing; // For Color
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Path to the output document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RichTextTable.docx");

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // -----------------------------------------------------------------
            // First row – simple header cells.
            // -----------------------------------------------------------------
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // -----------------------------------------------------------------
            // Second row – cells with rich text formatting using Runs.
            // -----------------------------------------------------------------
            // First cell with bold and red text.
            Cell cell1 = builder.InsertCell();
            Paragraph para1 = cell1.FirstParagraph;
            Run runBoldRed = new Run(doc, "Bold Red");
            runBoldRed.Font.Bold = true;
            runBoldRed.Font.Color = Color.Red; // Correct property
            para1.AppendChild(runBoldRed);
            // Append normal text after the formatted run.
            para1.AppendChild(new Run(doc, " normal text"));

            // Second cell with italic, blue text and underline.
            Cell cell2 = builder.InsertCell();
            Paragraph para2 = cell2.FirstParagraph;
            Run runItalicBlue = new Run(doc, "Italic Blue");
            runItalicBlue.Font.Italic = true;
            runItalicBlue.Font.Color = Color.Blue; // Correct property
            para2.AppendChild(runItalicBlue);
            Run runUnderline = new Run(doc, " underlined");
            runUnderline.Font.Underline = Underline.Single;
            para2.AppendChild(runUnderline);

            // Finish the row.
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new Exception($"Failed to create the output file at {outputPath}");
        }
    }
}
