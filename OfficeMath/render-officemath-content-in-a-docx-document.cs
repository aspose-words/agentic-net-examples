using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX document that contains OfficeMath objects.
        Document doc = new Document("OfficeMath.docx");

        // Retrieve the first OfficeMath node in the document.
        // The GetChild method searches the document tree for a node of the specified type.
        OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
        if (officeMath == null)
        {
            Console.WriteLine("No OfficeMath object found in the document.");
            return;
        }

        // Create ImageSaveOptions to control how the OfficeMath object is rendered.
        // Here we render the equation as a PNG image and increase the scale for better resolution.
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Scale = 5 // Render the equation five times its original size.
        };

        // Render the OfficeMath object to an image file.
        // GetMathRenderer creates an OfficeMathRenderer that can save the equation as an image.
        officeMath.GetMathRenderer().Save("RenderedOfficeMath.png", imgOptions);

        // Optionally, save the original document (or a modified version) back to DOCX.
        doc.Save("OfficeMath_Output.docx");
    }
}
