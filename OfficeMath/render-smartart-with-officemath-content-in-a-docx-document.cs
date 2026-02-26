using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Math;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load a DOCX that already contains SmartArt with OfficeMath objects.
        Document doc = new Document("SmartArtWithMath.docx");

        // Ensure that all SmartArt drawings are rendered using Aspose.Words' cold rendering engine.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            shape.UpdateSmartArtDrawing();
        }

        // Render each OfficeMath object to a separate PNG image.
        int mathIndex = 0;
        foreach (OfficeMath math in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            // Configure image rendering options (e.g., increase scale for higher resolution).
            ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Scale = 3 // render at three times the original size
            };

            string imagePath = $"OfficeMath_{mathIndex++}.png";
            math.GetMathRenderer().Save(imagePath, imgOptions);
        }

        // Save the document with updated SmartArt drawings.
        doc.Save("SmartArtWithMath_Updated.docx");
    }
}
