using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fonts;

class FontDemo
{
    static void Main()
    {
        // ---------- Create a new blank document ----------
        Document doc = new Document();                     // create
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Set font properties for the first line ----------
        builder.Font.Name = "Courier New";
        builder.Font.Size = 24;                           // points
        builder.Font.Color = Color.Blue;
        builder.Font.Bold = true;
        builder.Font.Italic = true;
        builder.Font.Underline = Underline.Double;
        builder.Font.HighlightColor = Color.Yellow;

        builder.Writeln("Formatted text using Font properties.");

        // ---------- Change font properties for the second line ----------
        builder.Font.Name = "Arial";
        builder.Font.Size = 18;
        builder.Font.Color = Color.DarkGreen;
        builder.Font.Bold = false;
        builder.Font.Italic = false;
        builder.Font.Underline = Underline.None;
        builder.Font.HighlightColor = Color.Transparent;

        builder.Writeln("Another line with different font settings.");

        // ---------- Save the document ----------
        string outPath = Path.Combine(Environment.CurrentDirectory, "FontDemo.docx");
        doc.Save(outPath);                                 // save

        // ---------- Load the saved document ----------
        Document loadedDoc = new Document(outPath);        // load

        // ---------- Modify the font of the first run ----------
        Paragraph firstPara = loadedDoc.FirstSection.Body.FirstParagraph;
        if (firstPara.Runs.Count > 0)
        {
            Font runFont = firstPara.Runs[0].Font;
            runFont.Name = "Times New Roman";
            runFont.Size = 28;
            runFont.Color = Color.Red;
        }

        // ---------- Save the modified document ----------
        string modifiedPath = Path.Combine(Environment.CurrentDirectory, "FontDemo_Modified.docx");
        loadedDoc.Save(modifiedPath);                     // save
    }
}
