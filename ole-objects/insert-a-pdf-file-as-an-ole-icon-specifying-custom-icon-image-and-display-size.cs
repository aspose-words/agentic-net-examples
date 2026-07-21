using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace OleObjectExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Resolve paths relative to the current working directory.
            string pdfPath = Path.GetFullPath("Sample.pdf");
            string iconPath = Path.GetFullPath("CustomIcon.ico");

            // Ensure the PDF file exists – create a minimal placeholder if it does not.
            if (!File.Exists(pdfPath))
            {
                // Very small valid PDF content.
                string minimalPdf = "%PDF-1.1\n%âãÏÓ\n1 0 obj\n<<>>\nendobj\nxref\n0 1\n0000000000 65535 f \ntrailer\n<<>>\nstartxref\n0\n%%EOF";
                File.WriteAllText(pdfPath, minimalPdf);
            }

            // If the custom icon file is missing, fall back to the default icon by passing null.
            if (!File.Exists(iconPath))
            {
                iconPath = null;
            }

            // Insert the PDF as an OLE object displayed as an icon.
            // Parameters:
            //   fileName   – full path to the PDF file.
            //   isLinked   – false = embed the file, true = link to the file.
            //   iconFile   – full path to the custom ICO file (null to use default).
            //   iconCaption– caption displayed under the icon (null to use file name).
            Shape oleShape = builder.InsertOleObjectAsIcon(pdfPath, false, iconPath, "Embedded PDF");

            // Set the display size of the icon (width and height in points).
            // Here we set the icon to 50 mm × 50 mm.
            oleShape.Width = ConvertUtil.MillimeterToPoint(50);
            oleShape.Height = ConvertUtil.MillimeterToPoint(50);

            // Save the resulting document.
            string outputPath = Path.GetFullPath("OleObjectWithPdfIcon.docx");
            doc.Save(outputPath);
        }
    }
}
