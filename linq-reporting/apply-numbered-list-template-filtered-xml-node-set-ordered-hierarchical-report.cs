using System;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Load the source XML file.
        XmlDocument xml = new XmlDocument();
        xml.Load("Data.xml"); // Path to your XML data.

        // Select only the nodes that satisfy the filter criteria.
        // Example: select <Item> elements with attribute type='A'.
        XmlNodeList nodes = xml.SelectNodes("//Item[@type='A']");

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list based on the default numbering template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Iterate through the filtered XML nodes and add them to the document as list items.
        foreach (XmlNode node in nodes)
        {
            // Determine the hierarchical level of the current node.
            // The depth of the node in the XML tree maps to the list level (0‑8).
            int level = GetNodeDepth(node) - 1; // Root element becomes level 0.
            if (level < 0) level = 0;
            if (level > 8) level = 8; // Word supports up to 9 levels (0‑8).

            // Apply the list and set the appropriate level.
            builder.ListFormat.List = numberedList;
            builder.ListFormat.ListLevelNumber = level;

            // Write the desired content of the node.
            // Here we prefer a <Title> child element; otherwise fall back to the node's inner text.
            string text = node.SelectSingleNode("Title")?.InnerText ?? node.InnerText;
            builder.Writeln(text);
        }

        // Finish the list.
        builder.ListFormat.RemoveNumbers();

        // Save the generated report.
        doc.Save("Report.docx");
    }

    // Helper method to calculate the depth of an XmlNode within its document.
    static int GetNodeDepth(XmlNode node)
    {
        int depth = 0;
        XmlNode current = node;
        while (current.ParentNode != null && current.ParentNode.NodeType != XmlNodeType.Document)
        {
            depth++;
            current = current.ParentNode;
        }
        return depth;
    }
}
