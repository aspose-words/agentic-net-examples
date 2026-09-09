using System;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // ---------- Create a sample image ----------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // ---------- Build a DOCX with the image and embed it in a custom XML part ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image into the document.
        Shape shape = builder.InsertImage(sampleImagePath);
        // Retrieve the image bytes from the shape.
        byte[] imageBytes = shape.ImageData.ToByteArray();
        // Encode the image as Base64 for storage in XML.
        string base64Image = Convert.ToBase64String(imageBytes);
        // Create XML that holds the image data and its original filename.
        string xmlContent = $"<images><image filename=\"sample.png\">{base64Image}</image></images>";
        // Add the custom XML part to the document.
        string partId = Guid.NewGuid().ToString("B");
        doc.CustomXmlParts.Add(partId, xmlContent);
        // Save the document.
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        doc.Save(docPath);

        // ---------- Load the document and extract images from custom XML parts ----------
        Document loadedDoc = new Document(docPath);
        int extractedCount = 0;

        foreach (CustomXmlPart part in loadedDoc.CustomXmlParts)
        {
            // The Data property contains the raw XML bytes.
            string partXml = Encoding.UTF8.GetString(part.Data);
            XDocument xDoc = XDocument.Parse(partXml);

            foreach (XElement imgElement in xDoc.Descendants("image"))
            {
                string fileName = (string)imgElement.Attribute("filename") ?? $"extracted_{extractedCount}.png";
                string base64Data = imgElement.Value;
                byte[] imgData = Convert.FromBase64String(base64Data);
                string outputPath = Path.Combine(artifactsDir, fileName);
                File.WriteAllBytes(outputPath, imgData);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from custom XML parts.");

        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to '{artifactsDir}'.");
    }
}
