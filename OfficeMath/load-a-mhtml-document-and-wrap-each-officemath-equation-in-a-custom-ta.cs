using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Paths to the input MHTML file and the output file.
        string inputPath = "input.mht";
        string outputPath = "output.mht";

        // Load the MHTML document. Enable conversion of shapes that contain EquationXML
        // to OfficeMath objects so that all equations are represented as OfficeMath nodes.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            ConvertShapeToOfficeMath = true
        };
        Document doc = new Document(inputPath, loadOptions);

        // Iterate over the OfficeMath nodes in reverse order because we will modify the
        // document tree while iterating.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        for (int i = officeMathNodes.Count - 1; i >= 0; i--)
        {
            OfficeMath om = (OfficeMath)officeMathNodes[i];

            // Insert the opening custom tag before the OfficeMath node.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(om);
            builder.InsertHtml("<custom>");

            // Insert the closing custom tag after the OfficeMath node.
            // If the OfficeMath node has a next sibling, move the builder there;
            // otherwise move it to the end of the parent node.
            Node next = om.NextSibling;
            if (next != null)
                builder.MoveTo(next);
            else
                builder.MoveTo(om.ParentNode.LastChild);
            builder.InsertHtml("</custom>");
        }

        // Save the modified document back to MHTML.
        // Export OfficeMath as MathML so that the equations remain visible in the HTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            OfficeMathOutputMode = HtmlOfficeMathOutputMode.MathML
        };
        doc.Save(outputPath, saveOptions);
    }
}
