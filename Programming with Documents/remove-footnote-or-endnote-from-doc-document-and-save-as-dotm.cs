// Load the existing DOC document
var doc = new Aspose.Words.Document("InputDocument.doc");

// Collect all footnote and endnote nodes in the document
var footnoteNodes = doc.GetChildNodes(Aspose.Words.NodeType.Footnote, true);

// Remove each footnote/endnote from its parent
foreach (Aspose.Words.Notes.Footnote footnote in footnoteNodes)
{
    footnote.Remove();
}

// Save the modified document as a DOTM template
doc.Save("OutputDocument.dotm", Aspose.Words.SaveFormat.Dotm);
