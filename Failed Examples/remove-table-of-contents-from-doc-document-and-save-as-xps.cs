// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveTocAndSaveAsXps
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("InputDocument.docx");

        // Find the first Table of Contents (if any) in the document and remove it.
        // NodeType.TableOfContents represents a TOC field.
        TableOfContents toc = (TableOfContents)doc.GetChild(NodeType.TableOfContents, 0, true);
        if (toc != null)
            toc.Remove();

        // Create XPS save options using the provided constructor.
        XpsSaveOptions saveOptions = new XpsSaveOptions();

        // Save the modified document as XPS using the save options.
        doc.Save("OutputDocument.xps", saveOptions);
    }
}
