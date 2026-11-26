using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

public class Program
{
    private static readonly XNamespace w = "http://schemas.openxml.org/wordprocessingml/2006/main";
    private static readonly XNamespace r = "http://schemas.openxml.org/officeDocument/2006/relationships";

    public static void Main()
    {
        // Folder containing the Word documents
        string folderPath = "Docs";

        // Ensure the folder exists (create if missing)
        if (!Directory.Exists(folderPath))
        {
            Directory.CreateDirectory(folderPath);
        }

        // URL of the online video to embed
        string videoUrl = "https://example.com/video.mp4";

        // Process each .docx file in the folder
        foreach (string filePath in Directory.GetFiles(folderPath, "*.docx"))
        {
            // Open the .docx package for update
            using (var archive = ZipFile.Open(filePath, ZipArchiveMode.Update))
            {
                // Load document.xml
                var docEntry = archive.GetEntry("word/document.xml");
                if (docEntry == null) continue;

                XDocument docXml;
                using (var stream = docEntry.Open())
                {
                    docXml = XDocument.Load(stream);
                }

                // Load or create document.xml.rels
                var relsEntry = archive.GetEntry("word/_rels/document.xml.rels");
                XDocument relsXml;
                if (relsEntry != null)
                {
                    using (var stream = relsEntry.Open())
                    {
                        relsXml = XDocument.Load(stream);
                    }
                }
                else
                {
                    // Create a new relationships document if missing
                    relsXml = new XDocument(
                        new XElement("Relationships")
                    );
                    relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                }

                // Determine a new relationship Id
                var existingIds = relsXml.Root
                    .Elements("Relationship")
                    .Attributes("Id")
                    .Select(a => a.Value)
                    .Where(id => id.StartsWith("rId"))
                    .Select(id => int.Parse(id.Substring(3)))
                    .DefaultIfEmpty(0);

                int nextIdNumber = existingIds.Max() + 1;
                string newRelId = $"rId{nextIdNumber}";

                // Add the hyperlink relationship
                var newRel = new XElement("Relationship",
                    new XAttribute("Id", newRelId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"),
                    new XAttribute("Target", videoUrl),
                    new XAttribute("TargetMode", "External")
                );
                relsXml.Root.Add(newRel);

                // Save the updated relationships file
                using (var stream = relsEntry.Open())
                {
                    stream.SetLength(0);
                    relsXml.Save(stream);
                }

                // Build the new paragraph with hyperlink
                var paragraph = new XElement(w + "p",
                    new XElement(w + "r",
                        new XElement(w + "t", "Embedded video: ")
                    ),
                    new XElement(w + "hyperlink",
                        new XAttribute(r + "id", newRelId),
                        new XElement(w + "r",
                            new XElement(w + "t", "Watch Video")
                        )
                    )
                );

                // Append the paragraph to the document body
                var body = docXml.Root.Element(w + "body");
                if (body != null)
                {
                    body.Add(paragraph);
                }

                // Save the updated document.xml
                using (var stream = docEntry.Open())
                {
                    stream.SetLength(0);
                    docXml.Save(stream);
                }
            }
        }
    }
}
