// Load an existing HTML document
var doc = new Aspose.Words.Document("input.html");

// Configure text watermark options (optional)
var watermarkOptions = new Aspose.Words.TextWatermarkOptions
{
    FontFamily = "Arial",
    FontSize = 36,
    Color = System.Drawing.Color.Gray,
    Layout = Aspose.Words.WatermarkLayout.Diagonal,
    IsSemitrasparent = true
};

// Add the text watermark to every page of the document
doc.Watermark.SetText("Confidential", watermarkOptions);

// Save the document back as HTML (watermark will be rendered as an image)
doc.Save("output.html");
