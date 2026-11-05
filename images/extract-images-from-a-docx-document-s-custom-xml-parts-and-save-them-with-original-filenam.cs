using System;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a deterministic sample image (sample.png)
        // -------------------------------------------------
        const string imagePath = "sample.png";
        const int width = 100;
        const int height = 100;

        // Create bitmap and draw a white background.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // Save the bitmap to a file.
        bitmap.Save(imagePath);
        // Clean up drawing objects.
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a new DOCX document and embed the image
        //    inside a custom XML part as a Base64 string.
        // -------------------------------------------------
        Document doc = new Document();

        byte[] imageBytes = File.ReadAllBytes(imagePath);
        string base64 = Convert.ToBase64String(imageBytes);
        string xmlContent = $"<images><image name=\"{Path.GetFileName(imagePath)}\" data=\"{base64}\"/></images>";

        // Add the custom XML part.
        doc.CustomXmlParts.Add("urn:customXml", xmlContent);

        // Save the document.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract images from the
        //    custom XML parts, saving them with their original
        //    filenames.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        int extractedCount = 0;

        foreach (var part in loadedDoc.CustomXmlParts)
        {
            // CustomXmlPart.Data returns a byte[]; convert it to a string.
            string data = Encoding.UTF8.GetString(part.Data);
            XDocument xdoc = XDocument.Parse(data);

            foreach (var imgElem in xdoc.Descendants("image"))
            {
                string name = (string)imgElem.Attribute("name");
                string dataAttr = (string)imgElem.Attribute("data");

                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(dataAttr))
                    continue;

                byte[] imgData = Convert.FromBase64String(dataAttr);
                File.WriteAllBytes(name, imgData);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from custom XML parts.");
    }
}
