using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class WatermarkDemo
{
    static void Main()
    {
        // User selection for watermark type.
        WatermarkType selectedType = GetUserSelectedWatermarkType();

        // Create a new blank document.
        Document doc = new Document();

        // Apply the appropriate watermark based on the selected type.
        switch (selectedType)
        {
            case WatermarkType.Text:
                // Create text watermark options (optional customization).
                TextWatermarkOptions textOptions = new TextWatermarkOptions
                {
                    FontFamily = "Arial",
                    FontSize = 36,
                    Color = Color.Gray,
                    Layout = WatermarkLayout.Diagonal,
                    IsSemitrasparent = false
                };
                doc.Watermark.SetText("Confidential", textOptions);
                break;

            case WatermarkType.Image:
                // Path to the image file to be used as a watermark.
                string imagePath = @"C:\Images\logo.png";

                // Create image watermark options (optional customization).
                ImageWatermarkOptions imageOptions = new ImageWatermarkOptions
                {
                    Scale = 5,          // Scale factor.
                    IsWashout = false   // Do not apply washout effect.
                };
                doc.Watermark.SetImage(imagePath, imageOptions);
                break;

            case WatermarkType.None:
                // No watermark is applied; ensure any existing watermark is removed.
                if (doc.Watermark.Type != WatermarkType.None)
                {
                    doc.Watermark.Remove();
                }
                break;

            default:
                throw new ArgumentOutOfRangeException(nameof(selectedType), "Unsupported watermark type.");
        }

        // Save the document with the applied watermark.
        doc.Save("WatermarkedDocument.docx");
    }

    // Mock method to simulate obtaining user selection.
    // In a real application this could be replaced with UI input, command‑line arguments, etc.
    static WatermarkType GetUserSelectedWatermarkType()
    {
        // For demonstration purposes, return Text watermark.
        // Change this return value to WatermarkType.Image or WatermarkType.None as needed.
        return WatermarkType.Text;
    }
}
