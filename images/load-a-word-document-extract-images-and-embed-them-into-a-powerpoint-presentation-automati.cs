using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample image file.
        const string imagePath = "input.png";
        const int imgWidth = 200;
        const int imgHeight = 200;
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black))
        {
            graphics.DrawRectangle(pen, 20, 20, imgWidth - 40, imgHeight - 40);
        }
        bitmap.Save(imagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Step 2: Create a Word document and insert the image.
        const string wordPath = "sample.docx";
        var wordDoc = new Document();
        var builder = new DocumentBuilder(wordDoc);
        builder.InsertImage(imagePath);
        wordDoc.Save(wordPath);

        // Step 3: Load the Word document and extract images.
        var loadedDoc = new Document(wordPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extractedImagePath = $"extracted-{extractedCount}.png";
                shape.ImageData.Save(extractedImagePath);
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the Word document.");

        // Step 4: Create a minimal PowerPoint (PPTX) file and embed the first extracted image.
        const string pptxPath = "output.pptx";
        string tempFolder = "pptx_temp";

        // Clean any previous temp folder.
        if (Directory.Exists(tempFolder))
            Directory.Delete(tempFolder, true);

        // Create required folder structure.
        Directory.CreateDirectory(Path.Combine(tempFolder, "_rels"));
        Directory.CreateDirectory(Path.Combine(tempFolder, "ppt"));
        Directory.CreateDirectory(Path.Combine(tempFolder, "ppt", "_rels"));
        Directory.CreateDirectory(Path.Combine(tempFolder, "ppt", "slides"));
        Directory.CreateDirectory(Path.Combine(tempFolder, "ppt", "slides", "_rels"));
        Directory.CreateDirectory(Path.Combine(tempFolder, "ppt", "media"));

        // Write [Content_Types].xml
        File.WriteAllText(Path.Combine(tempFolder, "[Content_Types].xml"),
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
    <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
    <Default Extension=""xml"" ContentType=""application/xml""/>
    <Default Extension=""png"" ContentType=""image/png""/>
    <Override PartName=""/ppt/presentation.xml"" ContentType=""application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml""/>
    <Override PartName=""/ppt/slides/slide1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.presentationml.slide+xml""/>
</Types>");

        // Write _rels/.rels
        File.WriteAllText(Path.Combine(tempFolder, "_rels", ".rels"),
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""ppt/presentation.xml""/>
</Relationships>");

        // Write ppt/presentation.xml
        File.WriteAllText(Path.Combine(tempFolder, "ppt", "presentation.xml"),
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<p:presentation xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main""
                xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
                xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <p:sldIdLst>
        <p:sldId id=""256"" r:id=""rId1""/>
    </p:sldIdLst>
    <p:sldMasterIdLst/>
    <p:notesMasterIdLst/>
    <p:handoutMasterIdLst/>
    <p:slideSize cx=""9144000"" cy=""6858000"" type=""screen4x3""/>
    <p:defaultTextStyle/>
</p:presentation>");

        // Write ppt/_rels/presentation.xml.rels
        File.WriteAllText(Path.Combine(tempFolder, "ppt", "_rels", "presentation.xml.rels"),
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"" Target=""slides/slide1.xml""/>
</Relationships>");

        // Write ppt/slides/slide1.xml (using the first extracted image)
        File.WriteAllText(Path.Combine(tempFolder, "ppt", "slides", "slide1.xml"),
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<p:sld xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main""
       xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
       xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id=""1"" name="""" />
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x=""0"" y=""0""/>
                    <a:ext cx=""0"" cy=""0""/>
                    <a:chOff x=""0"" y=""0""/>
                    <a:chExt cx=""0"" cy=""0""/>
                </a:xfrm>
            </p:grpSpPr>
            <p:pic>
                <p:nvPicPr>
                    <p:cNvPr id=""2"" name=""Picture 1""/>
                    <p:cNvPicPr/>
                    <p:nvPr/>
                </p:nvPicPr>
                <p:blipFill>
                    <a:blip r:embed=""rId2""/>
                    <a:stretch>
                        <a:fillRect/>
                    </a:stretch>
                </p:blipFill>
                <p:spPr>
                    <a:xfrm>
                        <a:off x=""1524000"" y=""1524000""/>
                        <a:ext cx=""4572000"" cy=""3429000""/>
                    </a:xfrm>
                    <a:prstGeom prst=""rect"">
                        <a:avLst/>
                    </a:prstGeom>
                </p:spPr>
            </p:pic>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>");

        // Write ppt/slides/_rels/slide1.xml.rels
        File.WriteAllText(Path.Combine(tempFolder, "ppt", "slides", "_rels", "slide1.xml.rels"),
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""../media/image1.png""/>
</Relationships>");

        // Copy the first extracted image into the media folder.
        string sourceImage = "extracted-0.png";
        string destImage = Path.Combine(tempFolder, "ppt", "media", "image1.png");
        File.Copy(sourceImage, destImage, true);

        // Create the PPTX zip package.
        if (File.Exists(pptxPath))
            File.Delete(pptxPath);
        ZipFile.CreateFromDirectory(tempFolder, pptxPath);

        // Clean up temporary folder.
        Directory.Delete(tempFolder, true);

        // Validation: ensure output file exists.
        if (!File.Exists(pptxPath))
            throw new InvalidOperationException("PowerPoint file was not created.");

        Console.WriteLine("Process completed successfully.");
    }
}
