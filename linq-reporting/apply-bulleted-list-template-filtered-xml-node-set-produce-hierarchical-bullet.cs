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

        // Sample XML data (replace with your own XML if desired).
        const string sampleXml = @"
<root>
    <item>Top level item 1</item>
    <item>
        <item>Second level item 1</item>
        <item>Second level item 2</item>
        <item>
            <item>Third level item</item>
        </item>
    </item>
    <item>Top level item 2</item>
</root>";

        // Load the source XML from the string.
        XmlDocument xml = new XmlDocument();
        xml.LoadXml(sampleXml);

        // Select all <item> elements.
        XmlNodeList xmlNodes = xml.SelectNodes("//item");

        foreach (XmlNode xmlNode in xmlNodes)
        {
            // Determine the hierarchical depth of the current node.
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
        doc.Save("output.docx");
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
