using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // HTML content containing MathML.
        const string html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
</head>
<body>
    <p>
        Here is an equation:
        <math xmlns='http://www.w3.org/1998/Math/MathML'>
            <mi>x</mi><mo>=</mo><mn>5</mn>
        </math>
    </p>
</body>
</html>";

        // Load the HTML from a memory stream with the appropriate options.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            ConvertShapeToOfficeMath = false
        };

        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
        {
            Document doc = new Document(stream, loadOptions);

            // Iterate through all Shape nodes (these represent the imported MathML equations).
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                // Position the shape inline with the surrounding text.
                shape.WrapType = WrapType.Inline;

                // Remove any vertical offset to keep the shape aligned with the text baseline.
                shape.Font.Position = 0;
            }

            // Save the modified document.
            doc.Save("Result.docx");
        }
    }
}
