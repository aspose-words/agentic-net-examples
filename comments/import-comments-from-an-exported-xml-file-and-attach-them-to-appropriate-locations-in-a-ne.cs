using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Aspose.Words;

public class CommentInfo
{
    public int ParagraphIndex { get; set; }
    public string Author { get; set; } = "";
    public string Initial { get; set; } = "";
    public DateTime Date { get; set; }
    public string Text { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the example files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "CommentImportExample");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample XML file that contains exported comment data.
        string xmlPath = Path.Combine(workDir, "comments.xml");
        CreateSampleCommentsXml(xmlPath);

        // 2. Load the comment data from the XML file.
        List<CommentInfo> commentInfos = LoadCommentsFromXml(xmlPath);

        // 3. Create a new blank document and add a few paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Store the created paragraphs so we can attach comments to them later.
        List<Paragraph> paragraphs = new List<Paragraph>();

        for (int i = 0; i < 3; i++)
        {
            builder.Writeln($"Paragraph {i + 1}: This is sample text for demonstration purposes.");
            Paragraph para = doc.FirstSection.Body.LastParagraph; // the paragraph just written
            paragraphs.Add(para);
        }

        // 4. Attach each imported comment to the appropriate paragraph.
        foreach (CommentInfo info in commentInfos)
        {
            // Guard against an invalid paragraph index.
            if (info.ParagraphIndex < 0 || info.ParagraphIndex >= paragraphs.Count)
                continue;

            Paragraph targetParagraph = paragraphs[info.ParagraphIndex];

            // Create the comment node with metadata.
            Comment comment = new Comment(doc, info.Author, info.Initial, info.Date);
            comment.SetText(info.Text);

            // Append the comment to the paragraph. This anchors the comment to the paragraph.
            targetParagraph.AppendChild(comment);
        }

        // 5. Save the resulting document.
        string outPath = Path.Combine(workDir, "CommentedDocument.docx");
        doc.Save(outPath);

        Console.WriteLine($"Document saved to: {outPath}");
    }

    // Creates a simple XML file with a few comment entries.
    private static void CreateSampleCommentsXml(string filePath)
    {
        XmlDocument xmlDoc = new XmlDocument();

        XmlElement root = xmlDoc.CreateElement("Comments");
        xmlDoc.AppendChild(root);

        // First comment attached to paragraph 0
        XmlElement comment1 = xmlDoc.CreateElement("Comment");
        comment1.SetAttribute("ParagraphIndex", "0");
        comment1.SetAttribute("Author", "Alice");
        comment1.SetAttribute("Initial", "A");
        comment1.SetAttribute("Date", DateTime.Now.AddDays(-1).ToString("o"));
        comment1.InnerText = "First comment on the first paragraph.";
        root.AppendChild(comment1);

        // Second comment attached to paragraph 2
        XmlElement comment2 = xmlDoc.CreateElement("Comment");
        comment2.SetAttribute("ParagraphIndex", "2");
        comment2.SetAttribute("Author", "Bob");
        comment2.SetAttribute("Initial", "B");
        comment2.SetAttribute("Date", DateTime.Now.ToString("o"));
        comment2.InnerText = "Another comment on the third paragraph.";
        root.AppendChild(comment2);

        xmlDoc.Save(filePath);
    }

    // Loads comment information from the XML file created above.
    private static List<CommentInfo> LoadCommentsFromXml(string filePath)
    {
        var list = new List<CommentInfo>();
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.Load(filePath);

        XmlNodeList nodes = xmlDoc.SelectNodes("/Comments/Comment");
        if (nodes == null) return list;

        foreach (XmlNode node in nodes)
        {
            if (node.Attributes == null) continue;

            var info = new CommentInfo();

            if (int.TryParse(node.Attributes["ParagraphIndex"]?.Value, out int idx))
                info.ParagraphIndex = idx;

            info.Author = node.Attributes["Author"]?.Value ?? "";
            info.Initial = node.Attributes["Initial"]?.Value ?? "";

            if (DateTime.TryParse(node.Attributes["Date"]?.Value, out DateTime dt))
                info.Date = dt;
            else
                info.Date = DateTime.Now;

            info.Text = node.InnerText ?? "";

            list.Add(info);
        }

        return list;
    }
}
