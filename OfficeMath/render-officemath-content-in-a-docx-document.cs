using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Saving;
using Aspose.Words.Rendering;

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
            Console.WriteLine("No OfficeMath objects were found in the document.");
            return;
        }

        // Create rendering options for the image.
        // Here we use PNG format and increase the scale to render a larger image.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Scale = 5 // Render the equation five times its original size.
        };

        // Render the OfficeMath object to an image file.
        // The GetMathRenderer method returns an OfficeMathRenderer that can save the equation.
        officeMath.GetMathRenderer().Save("RenderedOfficeMath.png", saveOptions);

        Console.WriteLine("OfficeMath has been rendered to 'RenderedOfficeMath.png'.");

        // OPTIONAL: Export the whole document to HTML with OfficeMath as MathML.
        // This demonstrates another way to render OfficeMath when saving the document.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            OfficeMathOutputMode = HtmlOfficeMathOutputMode.MathML
        };
        doc.Save("OfficeMath_AsMathML.html", htmlOptions);

        Console.WriteLine("Document saved as HTML with OfficeMath rendered as MathML.");
    }
}
