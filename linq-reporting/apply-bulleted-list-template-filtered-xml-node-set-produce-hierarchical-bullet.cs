using System;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Lists;

class XmlToBulletedList
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a multi‑level bulleted list using the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Load the source XML.
        XmlDocument xml = new XmlDocument();
        xml.Load("input.xml"); // replace with your XML file path

        // Select the nodes you want to turn into list items.
        // Example XPath selects all <item> elements; adjust as needed.
        XmlNodeList xmlNodes = xml.SelectNodes("//item");

        foreach (XmlNode xmlNode in xmlNodes)
        {
            // Determine the hierarchical depth of the current node.
            // Depth 0 => top‑level bullet, depth 1 => second level, etc.
            int depth = GetNodeDepth(xmlNode);

            // Ensure the depth does not exceed the maximum list level (0‑8).
            if (depth > 8) depth = 8;

            // Apply the list and set the appropriate level.
            builder.ListFormat.List = bulletList;
            builder.ListFormat.ListLevelNumber = depth;

            // Write the node's text as a list item.
            builder.Writeln(xmlNode.InnerText.Trim());
        }

        // End list formatting for any subsequent paragraphs.
        builder.ListFormat.List = null;

        // Save the resulting document.
        doc.Save("output.docx"); // replace with your desired output path
    }

    // Returns the number of element ancestors of the same node type.
    // This gives a simple hierarchy depth for XML structures.
    private static int GetNodeDepth(XmlNode node)
    {
        int depth = 0;
        XmlNode parent = node.ParentNode;
        while (parent != null && parent.NodeType == XmlNodeType.Element)
        {
            depth++;
            parent = parent.ParentNode;
        }
        return depth;
    }
}
