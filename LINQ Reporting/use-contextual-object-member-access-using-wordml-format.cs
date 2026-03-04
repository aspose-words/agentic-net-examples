using System;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Markup;

class WordmlContextualAccess
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert a StructuredDocumentTag (content control) into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        sdt.Title = "SampleTag";
        sdt.Tag = "SampleTagId";

        // Insert the content control at the current cursor position.
        builder.InsertNode(sdt);
        builder.Writeln("Content inside the tag.");

        // Retrieve the WORDML (Flat OPC) representation of the StructuredDocumentTag.
        string wordml = sdt.WordOpenXML;

        // Load the WORDML into an XmlDocument for XPath processing.
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(wordml);

        // Define an XPath that selects the Title attribute of the StructuredDocumentTag.
        // The WORDML uses the "w" namespace for WordprocessingML elements.
        XmlNamespaceManager nsMgr = new XmlNamespaceManager(xmlDoc.NameTable);
        nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        // Select the Title attribute node.
        XmlNode titleNode = xmlDoc.SelectSingleNode("//w:sdtPr/w:alias/@w:val", nsMgr);

        // Output the original title value.
        Console.WriteLine("Original Title: " + (titleNode?.Value ?? "Not found"));

        // Modify the title attribute directly in the WORDML.
        if (titleNode != null)
        {
            titleNode.Value = "UpdatedTagTitle";
        }

        // Save the modified WORDML back to the StructuredDocumentTag.
        // StructuredDocumentTag does not provide a setter for WordOpenXML, so we replace the original tag
        // with a new one created from the modified XML.
        // Load the modified XML into a temporary document.
        Document tempDoc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(xmlDoc.OuterXml)));
        StructuredDocumentTag newSdt = tempDoc.GetChild(NodeType.StructuredDocumentTag, 0, true) as StructuredDocumentTag;
        if (newSdt == null)
        {
            Console.WriteLine("Failed to load modified StructuredDocumentTag from WORDML.");
            return;
        }

        // Import the new tag into the original document.
        StructuredDocumentTag updatedSdt = (StructuredDocumentTag)doc.ImportNode(newSdt, true);

        // Replace the old tag with the updated one in the document tree.
        CompositeNode parent = sdt.ParentNode as CompositeNode;
        if (parent != null)
        {
            parent.InsertAfter(updatedSdt, sdt);
            sdt.Remove();
        }
        else
        {
            Console.WriteLine("Parent node is not a CompositeNode; cannot replace the tag.");
            return;
        }

        // Save the document to a file.
        doc.Save("WordmlContextualAccess.docx");
    }
}
