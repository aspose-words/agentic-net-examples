using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

class InsertDotIntoContentControl
{
    static void Main()
    {
        // Path to the DOT template that will be inserted.
        string dotTemplatePath = @"C:\Templates\SampleTemplate.dot";

        // Create a new blank document.
        Document mainDoc = new Document();

        // Create a DocumentBuilder for the new document.
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Insert a rich‑text content control (structured document tag) where the template will go.
        // SdtType.RichText creates a plain content control that can hold any block level content.
        StructuredDocumentTag sdt = new StructuredDocumentTag(mainDoc, SdtType.RichText, MarkupLevel.Block);
        sdt.Title = "InsertedTemplate";
        sdt.Tag = "TemplateTag";

        // Append the content control to the document body.
        builder.CurrentParagraph.AppendChild(sdt);

        // Load the DOT template document.
        Document templateDoc = new Document(dotTemplatePath);

        // Prepare a NodeImporter to efficiently import nodes from the template into the main document.
        NodeImporter importer = new NodeImporter(templateDoc, mainDoc, ImportFormatMode.KeepSourceFormatting);

        // Import each section of the template into the content control.
        // The content control can contain block nodes, so we import the body of each section.
        foreach (Section srcSection in templateDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the final empty paragraph that Word adds to each section.
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                // Import the node into the destination document.
                Node importedNode = importer.ImportNode(srcNode, true);

                // Add the imported node to the content control.
                sdt.AppendChild(importedNode);
            }
        }

        // Save the resulting document as a DOT file.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);
        mainDoc.Save(@"C:\Output\ResultDocument.dot", saveOptions);
    }
}
