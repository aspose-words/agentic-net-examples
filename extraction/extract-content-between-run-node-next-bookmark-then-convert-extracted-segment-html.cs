using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractRunToBookmarkHtml
{
    static void Main()
    {
        // Load the source document (lifecycle rule: load)
        Document srcDoc = new Document("Input.docx");

        // Find the first Run node in the document.
        Run startRun = srcDoc.GetChildNodes(NodeType.Run, true)[0] as Run;
        if (startRun == null)
        {
            Console.WriteLine("No Run node found.");
            return;
        }

        // Locate the next BookmarkStart after the Run.
        BookmarkStart nextBookmark = null;
        Node node = startRun.NextSibling;
        while (node != null)
        {
            if (node.NodeType == NodeType.BookmarkStart)
            {
                nextBookmark = (BookmarkStart)node;
                break;
            }
            node = node.NextSibling;
        }

        if (nextBookmark == null)
        {
            Console.WriteLine("No following bookmark found.");
            return;
        }

        // Create an empty document to hold the extracted segment (lifecycle rule: create)
        Document fragmentDoc = new Document();

        // Ensure the fragment has at least one section and body.
        if (fragmentDoc.FirstSection == null)
            fragmentDoc.AppendChild(new Section(fragmentDoc));
        if (fragmentDoc.FirstSection.Body == null)
            fragmentDoc.FirstSection.AppendChild(new Body(fragmentDoc));

        // Import nodes from the source document into the fragment until the bookmark is reached.
        NodeImporter importer = new NodeImporter(srcDoc, fragmentDoc, ImportFormatMode.KeepSourceFormatting);
        Node cur = startRun;
        while (cur != null && cur != nextBookmark)
        {
            Node imported = importer.ImportNode(cur, true);
            fragmentDoc.FirstSection.Body.AppendChild(imported);
            cur = cur.NextSibling;
        }

        // Convert the extracted fragment to HTML.
        string html = fragmentDoc.ToString(SaveFormat.Html);

        // Output the HTML to console.
        Console.WriteLine("Extracted HTML:");
        Console.WriteLine(html);

        // Optionally, save the HTML to a file (lifecycle rule: save)
        File.WriteAllText("ExtractedSegment.html", html);
    }
}
