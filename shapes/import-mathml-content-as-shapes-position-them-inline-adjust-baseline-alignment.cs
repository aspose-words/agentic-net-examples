using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an HTML file that contains MathML. The MathML will be imported as Shape objects.
        // Setting ConvertShapeToOfficeMath to false ensures the MathML stays as shapes rather than being converted to OfficeMath.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.ConvertShapeToOfficeMath = false;

        Document doc = new Document("MathML.html", loadOptions);

        // Iterate through all Shape nodes (these represent the imported MathML equations).
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            // Position the shape inline with the surrounding text.
            shape.WrapType = WrapType.Inline;

            // In older Aspose.Words versions the Font class does not expose BaselineAlignment.
            // To keep the shape aligned with the text baseline we can set the vertical offset to 0.
            // This mimics the default baseline alignment.
            shape.Font.Position = 0; // optional – removes any vertical offset
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
