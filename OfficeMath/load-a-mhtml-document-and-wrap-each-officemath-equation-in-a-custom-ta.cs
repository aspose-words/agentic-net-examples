using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Math; // <-- required for OfficeMath

class WrapOfficeMathInMhtml
{
    static void Main()
    {
        // Load the MHTML document. Use HtmlLoadOptions to enable conversion of shapes with EquationXML to OfficeMath.
        var loadOptions = new HtmlLoadOptions
        {
            ConvertShapeToOfficeMath = true
        };
        Document doc = new Document("input.mht", loadOptions);

        // Get all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate backwards so that inserting nodes does not affect the collection indexing.
        for (int i = officeMathNodes.Count - 1; i >= 0; i--)
        {
            OfficeMath officeMath = (OfficeMath)officeMathNodes[i];

            // Insert an opening custom tag before the OfficeMath node.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(officeMath);               // insertion point is *before* the node
            builder.InsertHtml("<my-equation>");

            // Insert a closing custom tag after the OfficeMath node.
            // After the opening tag is inserted the OfficeMath node itself is unchanged, so we move to the node that follows it.
            Node nextNode = officeMath.NextSibling;
            if (nextNode != null)
            {
                builder.MoveTo(nextNode);            // insertion point is before the next sibling → after the OfficeMath node
                builder.InsertHtml("</my-equation>");
            }
            else
            {
                // If the OfficeMath node is the last child, move to its parent and insert at the end of the parent.
                builder.MoveTo(officeMath.ParentNode);
                builder.InsertHtml("</my-equation>");
            }
        }

        // Save the modified document back to MHTML.
        doc.Save("output.mht", SaveFormat.Mhtml);
    }
}
