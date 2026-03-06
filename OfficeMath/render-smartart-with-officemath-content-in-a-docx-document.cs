using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class RenderSmartArtWithOfficeMath
{
    static void Main()
    {
        // Load an existing DOCX that contains SmartArt with OfficeMath objects.
        Document doc = new Document("SmartArtOfficeMath.docx");

        // Iterate through all shapes in the document.
        // The UpdateSmartArtDrawing method will refresh the pre‑rendered drawing
        // for SmartArt shapes, invoking Aspose.Words's cold rendering engine.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            shape.UpdateSmartArtDrawing();
        }

        // Save the document. The SmartArt now contains correctly rendered OfficeMath content.
        doc.Save("SmartArtOfficeMath_Rendered.docx");
    }
}
