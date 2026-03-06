using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddTextWatermarkToTxt
{
    static void Main()
    {
        // Path to the folder containing the input TXT file.
        string dataDir = @"C:\Data\";

        // Load the TXT file into an Aspose.Words Document.
        Document doc = new Document(dataDir + "input.txt");

        // Configure watermark appearance.
        TextWatermarkOptions options = new TextWatermarkOptions();
        options.FontFamily = "Arial";          // Font family.
        options.FontSize = 36;                 // Font size.
        options.Color = Color.Black;           // Font color.
        options.Layout = WatermarkLayout.Diagonal; // Diagonal layout.
        options.IsSemitrasparent = false;      // Opaque watermark.

        // Add the text watermark to the document.
        doc.Watermark.SetText("Confidential", options);

        // Save the document (watermark is supported in formats like DOCX, PDF, etc.).
        doc.Save(dataDir + "output.docx");
    }
}
