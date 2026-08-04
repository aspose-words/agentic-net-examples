using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ---------- First cell ----------
        // Insert a cell and add a paragraph with formatted runs.
        Cell cell1 = builder.InsertCell();
        Paragraph para1 = cell1.FirstParagraph;

        Run runBold = new Run(doc, "Bold");
        runBold.Font.Bold = true;
        para1.AppendChild(runBold);

        Run runSpace1 = new Run(doc, " ");
        para1.AppendChild(runSpace1);

        Run runItalic = new Run(doc, "Italic");
        runItalic.Font.Italic = true;
        para1.AppendChild(runItalic);

        Run runSpace2 = new Run(doc, " ");
        para1.AppendChild(runSpace2);

        Run runRed = new Run(doc, "Red");
        runRed.Font.Color = Color.Red;
        para1.AppendChild(runRed);

        // ---------- Second cell ----------
        Cell cell2 = builder.InsertCell();
        Paragraph para2 = cell2.FirstParagraph;

        Run runUnderline = new Run(doc, "Underline");
        runUnderline.Font.Underline = Underline.Single;
        para2.AppendChild(runUnderline);

        Run runSpace3 = new Run(doc, " ");
        para2.AppendChild(runSpace3);

        Run runBlue = new Run(doc, "Blue");
        runBlue.Font.Color = Color.Blue;
        para2.AppendChild(runBlue);

        Run runSpace4 = new Run(doc, " ");
        para2.AppendChild(runSpace4);

        Run runLarge = new Run(doc, "Large");
        runLarge.Font.Size = 16;
        para2.AppendChild(runLarge);

        // End the current row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RichTextTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
