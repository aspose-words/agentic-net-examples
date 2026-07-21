using System;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Drawing;

namespace AsposeWordsImageExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Prepare deterministic paths.
            string workingDir = Directory.GetCurrentDirectory();
            string imagePath = Path.Combine(workingDir, "sample.png");
            string docPath = Path.Combine(workingDir, "sample.docx");

            // -------------------------------------------------
            // Step 1: Create a sample image using Aspose.Drawing.
            // -------------------------------------------------
            const int imgWidth = 100;
            const int imgHeight = 100;
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            graphics.Clear(Aspose.Drawing.Color.White);
            bitmap.Save(imagePath);
            graphics.Dispose();
            bitmap.Dispose();

            // -------------------------------------------------
            // Step 2: Embed the image (as Base64) into a custom XML part.
            // -------------------------------------------------
            byte[] imageBytes = File.ReadAllBytes(imagePath);
            string base64Image = Convert.ToBase64String(imageBytes);
            string xmlContent = $"<images><image name=\"{Path.GetFileName(imagePath)}\">{base64Image}</image></images>";

            Document doc = new Document();
            string partId = Guid.NewGuid().ToString("B");
            // Add the custom XML part (data is UTF‑8 encoded XML). Use the string overload.
            doc.CustomXmlParts.Add(partId, xmlContent);

            // Save the document that now contains the custom XML part.
            doc.Save(docPath);

            // -------------------------------------------------
            // Step 3: Load the document and extract images from its custom XML parts.
            // -------------------------------------------------
            Document loadedDoc = new Document(docPath);
            int extractedCount = 0;

            foreach (CustomXmlPart part in loadedDoc.CustomXmlParts)
            {
                // Convert the part's binary data back to a string.
                string partXml = Encoding.UTF8.GetString(part.Data);
                XDocument xDoc = XDocument.Parse(partXml);

                foreach (XElement imgElement in xDoc.Descendants("image"))
                {
                    XAttribute nameAttr = imgElement.Attribute("name");
                    if (nameAttr == null) continue;

                    string fileName = nameAttr.Value;
                    string base64Data = imgElement.Value.Trim();
                    if (string.IsNullOrEmpty(base64Data)) continue;

                    byte[] imgData = Convert.FromBase64String(base64Data);
                    string outputPath = Path.Combine(workingDir, fileName);
                    File.WriteAllBytes(outputPath, imgData);
                    extractedCount++;
                }
            }

            // -------------------------------------------------
            // Validation: ensure at least one image was extracted.
            // -------------------------------------------------
            if (extractedCount == 0)
                throw new InvalidOperationException("No images were extracted from the custom XML parts.");

            // The example finishes without requiring user interaction.
        }
    }
}
