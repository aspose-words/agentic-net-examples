// Load the source DOCX file
var sourceDoc = new Aspose.Words.Document("Input.docx");

// Define the start and end bookmarks that delimit the region to extract
var startBookmark = sourceDoc.Range.Bookmarks["Start"];
var endBookmark = sourceDoc.Range.Bookmarks["End"];

// Ensure both bookmarks exist
if (startBookmark == null || endBookmark == null)
    throw new System.Exception("Required bookmarks not found.");

// Collect the text of all nodes that lie between the two bookmark markers
var extractedText = new System.Text.StringBuilder();
var currentNode = startBookmark.BookmarkStart.NextSibling;

// Traverse nodes until we reach the end bookmark start node
while (currentNode != null && currentNode != endBookmark.BookmarkStart)
{
    // Append the text of the current node (including its children)
    extractedText.Append(currentNode.GetText());
    currentNode = currentNode.NextSibling;
}

// Create a new blank document to hold the extracted content
var resultDoc = new Aspose.Words.Document();
var builder = new Aspose.Words.DocumentBuilder(resultDoc);

// Write the extracted text into the new document
builder.Writeln(extractedText.ToString().Trim());

// Save the result as a plain‑text file (extension determines format)
resultDoc.Save("ExtractedContent.txt");
