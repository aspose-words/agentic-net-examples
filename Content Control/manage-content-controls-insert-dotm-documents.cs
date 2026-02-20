using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOTM template that we want to insert.
        Document sourceTemplate = new Document(@"C:\Templates\MyTemplate.dotm");

        // Create a new blank document that will receive the content.
        Document destination = new Document();

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);

        // ------------------------------------------------------------
        // Create a rich‑text content control (structured document tag).
        // The DocumentBuilder class does not expose an InsertContentControl
        // method in older/standard Aspose.Words versions, so we instantiate
        // the StructuredDocumentTag directly and insert it into the document.
        // ------------------------------------------------------------
        StructuredDocumentTag sdt = new StructuredDocumentTag(
            destination,               // the owner document
            SdtType.RichText,          // type of the content control
            MarkupLevel.Block);        // block‑level control (behaves like a paragraph)

        // Insert the content control at the current cursor position.
        builder.InsertNode(sdt);
        // Move the builder inside the newly created control so that any further
        // inserted nodes become children of the control.
        builder.MoveTo(sdt);

        sdt.Title = "InsertedTemplate";
        sdt.Tag = "TemplateTag";

        // Prepare a NodeImporter to copy nodes from the source template to the destination document.
        NodeImporter importer = new NodeImporter(sourceTemplate, destination, ImportFormatMode.KeepSourceFormatting);

        // Import all nodes from the body of the first section of the source template
        // and add them as children of the content control.
        foreach (Node node in sourceTemplate.FirstSection.Body)
        {
            // Skip the final empty paragraph that Word adds to each section.
            if (node.NodeType == NodeType.Paragraph)
            {
                Paragraph para = (Paragraph)node;
                if (para.IsEndOfSection && !para.HasChildNodes)
                    continue;
            }

            Node importedNode = importer.ImportNode(node, true);
            sdt.AppendChild(importedNode);
        }

        // Save the resulting document. Use OoxmlSaveOptions to specify the format explicitly.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        destination.Save(@"C:\Output\DocumentWithInsertedTemplate.docx", saveOptions);
    }
}
