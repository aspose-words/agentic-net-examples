using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

// Load the existing HTML document.
string inputPath = @"C:\Docs\input.html";
Document doc = new Document(inputPath);

// Configure the text watermark appearance.
TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
{
    FontFamily = "Arial",
    FontSize = 36,
    Color = Color.Gray,
    Layout = WatermarkLayout.Diagonal,
    IsSemitrasparent = false
};

// Add the text watermark to the document.
doc.Watermark.SetText("Confidential", watermarkOptions);

// Save the document back to HTML format.
string outputPath = @"C:\Docs\output.html";
doc.Save(outputPath, SaveOptions.CreateSaveOptions(SaveFormat.Html));
