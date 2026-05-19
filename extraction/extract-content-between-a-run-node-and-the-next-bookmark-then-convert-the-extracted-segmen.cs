using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document containing a Run followed by some text
        //    and then a bookmark.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before the target run.");
        builder.Font.Bold = false;
        builder.Font.Italic = false;
        builder.Write("StartRun");                     // the run we will locate
        builder.Write("MiddleText");                   // text that should be extracted
        builder.StartBookmark("NextBookmark");        // bookmark placed after the text
        builder.EndBookmark("NextBookmark");
        builder.Writeln(); // move to next paragraph
        builder.Writeln("Paragraph after the bookmark.");

        const string sourcePath = "sample.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Locate the Run node with the exact text "StartRun".
        // -----------------------------------------------------------------
        Run startRun = sourceDoc.GetChildNodes(NodeType.Run, true)
            .Cast<Run>()
            .FirstOrDefault(r => r.Text == "StartRun");

        if (startRun == null)
            throw new InvalidOperationException("The start Run node was not found.");

        // -----------------------------------------------------------------
        // 3. Find the next BookmarkStart node after the identified Run.
        // -----------------------------------------------------------------
        Node nextBookmarkStart = null;
        for (Node node = startRun.NextSibling; node != null; node = node.NextSibling)
        {
            if (node.NodeType == NodeType.BookmarkStart)
            {
                nextBookmarkStart = node;
                break;
            }
        }

        if (nextBookmarkStart == null)
            throw new InvalidOperationException("No subsequent bookmark was found after the start Run.");

        // -----------------------------------------------------------------
        // 4. Extract the textual content that lies between the Run and the bookmark.
        // -----------------------------------------------------------------
        string extractedText = string.Empty;
        for (Node node = startRun.NextSibling; node != null && node != nextBookmarkStart; node = node.NextSibling)
        {
            // For Run nodes we take the raw Text; for other nodes we use GetText().
            if (node.NodeType == NodeType.Run)
                extractedText += ((Run)node).Text;
            else
                extractedText += node.GetText();
        }

        if (string.IsNullOrEmpty(extractedText))
            throw new InvalidOperationException("No text was extracted between the Run and the bookmark.");

        // -----------------------------------------------------------------
        // 5. Create a new document that contains the extracted segment.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        // Remove the default empty section/body created by the constructor.
        resultDoc.RemoveAllChildren();

        Section section = new Section(resultDoc);
        resultDoc.AppendChild(section);
        Body body = new Body(resultDoc);
        section.AppendChild(body);

        Paragraph para = new Paragraph(resultDoc);
        para.AppendChild(new Run(resultDoc, extractedText));
        body.AppendChild(para);

        // -----------------------------------------------------------------
        // 6. Save the extracted segment as HTML.
        // -----------------------------------------------------------------
        const string htmlPath = "extracted.html";
        resultDoc.Save(htmlPath, SaveFormat.Html);

        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // -----------------------------------------------------------------
        // 7. Write a small JSON report about the operation.
        // -----------------------------------------------------------------
        var report = new
        {
            SourceDocument = sourcePath,
            ExtractedText = extractedText,
            HtmlOutput = htmlPath,
            ExtractionTimestamp = DateTime.UtcNow
        };

        File.WriteAllText("extraction_report.json",
            JsonConvert.SerializeObject(report, Formatting.Indented));
    }
}
