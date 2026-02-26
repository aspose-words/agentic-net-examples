// Load a PDF file as an Aspose.Words document
string pdfPath = @"C:\Input\sample.pdf";
Aspose.Words.Document sourceDoc = new Aspose.Words.Document(pdfPath);

// Assume the PDF (now a Word document) contains two bookmarks named "Start" and "End"
// that mark the beginning and the end of the region we want to extract.
Aspose.Words.Bookmark startBookmark = sourceDoc.Range.Bookmarks["Start"];
Aspose.Words.Bookmark endBookmark = sourceDoc.Range.Bookmarks["End"];

// Validate that both bookmarks exist
if (startBookmark == null || endBookmark == null)
    throw new InvalidOperationException("Required bookmarks 'Start' and/or 'End' were not found.");

// The content between the two bookmarks can be obtained by iterating the nodes
// that lie after the start bookmark and before the end bookmark.
Aspose.Words.Node startNode = startBookmark.BookmarkStart;
Aspose.Words.Node endNode = endBookmark.BookmarkEnd;

// Create a new empty document that will hold the extracted content.
Aspose.Words.Document extractedDoc = new Aspose.Words.Document();
Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(extractedDoc);

// Move through the sibling nodes starting after the start bookmark.
Aspose.Words.Node current = startNode.NextSibling;
while (current != null && current != endNode)
{
    // Import the node into the new document to preserve formatting.
    Aspose.Words.Node imported = extractedDoc.ImportNode(current, true);
    builder.InsertNode(imported);
    current = current.NextSibling;
}

// Save the extracted portion as a separate DOCX file.
string outputPath = @"C:\Output\extracted_content.docx";
extractedDoc.Save(outputPath);
