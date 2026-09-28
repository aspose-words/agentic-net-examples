using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

public class Program
{
    public static void Main()
    {
        // Paths (files will be created in the executable's working directory)
        string inputDocx = "input.docx";
        string outputDocx = "output.docx";
        string newImage = "newImage.png";

        // Ensure a tiny PNG exists (1x1 transparent pixel)
        if (!File.Exists(newImage))
        {
            const string pngBase64 =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(pngBase64);
            File.WriteAllBytes(newImage, pngBytes);
        }

        // Ensure a minimal DOCX exists with a placeholder OLE object
        if (!File.Exists(inputDocx))
        {
            using (var zip = ZipFile.Open(inputDocx, ZipArchiveMode.Create))
            {
                // [Content_Types].xml
                var ctEntry = zip.CreateEntry("[Content_Types].xml");
                using (var ctStream = ctEntry.Open())
                {
                    XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
                    XDocument ctDoc = new XDocument(
                        new XElement(ns + "Types",
                            new XElement(ns + "Default",
                                new XAttribute("Extension", "rels"),
                                new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                            new XElement(ns + "Default",
                                new XAttribute("Extension", "xml"),
                                new XAttribute("ContentType", "application/xml")),
                            new XElement(ns + "Override",
                                new XAttribute("PartName", "/word/embeddings/ole.bin"),
                                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.oleObject"))
                        )
                    );
                    ctDoc.Save(ctStream);
                }

                // Placeholder OLE object (empty binary)
                var oleEntry = zip.CreateEntry("word/embeddings/ole.bin");
                using (var oleStream = oleEntry.Open())
                {
                    // Write a few bytes so the entry is not zero‑length
                    oleStream.Write(new byte[] { 0x00, 0x01, 0x02, 0x03 }, 0, 4);
                }
            }
        }

        // Copy the original document to a new file that will be modified
        File.Copy(inputDocx, outputDocx, true);

        // Open the DOCX (ZIP archive) for updating
        using (var archive = ZipFile.Open(outputDocx, ZipArchiveMode.Update))
        {
            // Locate the first OLE object part inside word/embeddings/
            var oleEntry = archive.Entries
                .FirstOrDefault(e => e.FullName.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase));

            if (oleEntry != null && File.Exists(newImage))
            {
                // Replace the OLE part content with the new image bytes
                using (var entryStream = oleEntry.Open())
                using (var imageStream = File.OpenRead(newImage))
                {
                    // Overwrite the existing content
                    entryStream.SetLength(0);
                    imageStream.CopyTo(entryStream);
                }

                // Update the content type for this part to image/png
                var ctEntry = archive.GetEntry("[Content_Types].xml");
                if (ctEntry != null)
                {
                    XDocument ctDoc;
                    using (var ctStream = ctEntry.Open())
                    {
                        ctDoc = XDocument.Load(ctStream);
                    }

                    XNamespace ns = ctDoc.Root.GetDefaultNamespace();

                    // Find the Override element for the OLE part
                    var overrideElem = ctDoc.Root
                        .Elements(ns + "Override")
                        .FirstOrDefault(x => (string)x.Attribute("PartName") == "/" + oleEntry.FullName);

                    if (overrideElem != null)
                    {
                        overrideElem.SetAttributeValue("ContentType", "image/png");
                    }
                    else
                    {
                        // If not found, add a new Override element
                        ctDoc.Root.Add(new XElement(ns + "Override",
                            new XAttribute("PartName", "/" + oleEntry.FullName),
                            new XAttribute("ContentType", "image/png")));
                    }

                    // Save the modified [Content_Types].xml back into the archive
                    using (var ctWriteStream = ctEntry.Open())
                    {
                        ctWriteStream.SetLength(0);
                        ctDoc.Save(ctWriteStream);
                    }
                }
            }
        }

        // At this point, output.docx contains the replaced OLE object (now a PNG image).
        // No further action required; the program exits automatically.
    }
}
