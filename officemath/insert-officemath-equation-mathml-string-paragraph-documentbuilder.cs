using System;
using Aspose.Words;
using Aspose.Words.Math;

class InsertOfficeMathFromMathML
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // MathML string that represents the equation to be inserted.
        // Example: quadratic formula.
        string mathMl = @"
            <math xmlns='http://www.w3.org/1998/Math/MathML'>
                <mi>x</mi>
                <mo>=</mo>
                <mfrac>
                    <mrow>
                        <mo>-</mo><mi>b</mi>
                        <mo>±</mo>
                        <msqrt>
                            <msup><mi>b</mi><mn>2</mn></msup>
                            <mo>-</mo>
                            <mn>4</mn><mi>a</mi><mi>c</mi>
                        </msqrt>
                    </mrow>
                    <mrow><mn>2</mn><mi>a</mi></mrow>
                </mfrac>
            </math>";

        // Insert the MathML as HTML. Aspose.Words parses MathML and creates an OfficeMath node.
        builder.InsertHtml(mathMl);

        // Optionally, retrieve the inserted OfficeMath node to adjust its display type.
        OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
        if (officeMath != null)
        {
            // Display the equation on its own line.
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            // Align the equation to the left.
            officeMath.Justification = OfficeMathJustification.Left;
        }

        // Save the document to a .docx file.
        doc.Save("OfficeMathFromMathML.docx");
    }
}
