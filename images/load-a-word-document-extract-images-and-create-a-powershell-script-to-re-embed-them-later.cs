using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

namespace ImageExtractionExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare directories and file names
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
            string documentPath = Path.Combine(artifactsDir, "sample.docx");
            string extractedImagesDir = Path.Combine(artifactsDir, "ExtractedImages");
            Directory.CreateDirectory(extractedImagesDir);
            string powershellScriptPath = Path.Combine(artifactsDir, "ReembedImages.ps1");

            // 1. Create a deterministic sample image using Aspose.Drawing
            CreateSampleImage(sampleImagePath);

            // 2. Create a Word document and insert the sample image
            CreateDocumentWithImage(documentPath, sampleImagePath);

            // 3. Load the document and extract all images
            int extractedCount = ExtractImages(documentPath, extractedImagesDir);
            if (extractedCount == 0)
                throw new InvalidOperationException("No images were extracted from the document.");

            // 4. Generate a PowerShell script that re‑embeds the extracted images
            GeneratePowerShellScript(documentPath, extractedImagesDir, powershellScriptPath);
            if (!File.Exists(powershellScriptPath))
                throw new InvalidOperationException("Failed to create the PowerShell script.");
        }

        // Creates a simple 200x200 white PNG with a black rectangle.
        private static void CreateSampleImage(string filePath)
        {
            const int width = 200;
            const int height = 200;

            using (Bitmap bitmap = new Bitmap(width, height))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }

                bitmap.Save(filePath);
            }
        }

        // Creates a new Word document and inserts the image at the end of the first paragraph.
        private static void CreateDocumentWithImage(string docPath, string imagePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Document with an embedded image:");
            builder.InsertImage(imagePath);
            doc.Save(docPath);
        }

        // Extracts all images from the document and saves them to the target folder.
        // Returns the number of extracted images.
        private static int ExtractImages(string docPath, string outputFolder)
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = $"extracted_{imageIndex}{extension}";
                string fullPath = Path.Combine(outputFolder, fileName);
                shape.ImageData.Save(fullPath);
                imageIndex++;
            }

            return imageIndex;
        }

        // Generates a simple PowerShell script that demonstrates how to re‑embed the extracted images.
        private static void GeneratePowerShellScript(string docPath, string imagesFolder, string scriptPath)
        {
            var sb = new StringBuilder();

            sb.AppendLine("# PowerShell script to re‑embed extracted images into a Word document");
            sb.AppendLine("$docPath = \"" + EscapeForPowerShell(docPath) + "\"");
            sb.AppendLine("$imagesFolder = \"" + EscapeForPowerShell(imagesFolder) + "\"");
            sb.AppendLine("$outputPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($docPath), \"Reembedded.docx\")");
            sb.AppendLine();
            sb.AppendLine("Add-Type -Path \"$(Join-Path $PSScriptRoot \"Aspose.Words.dll\")\"");
            sb.AppendLine("Add-Type -Path \"$(Join-Path $PSScriptRoot \"Aspose.Drawing.Common.dll\")\"");
            sb.AppendLine();
            sb.AppendLine("$doc = New-Object Aspose.Words.Document $docPath");
            sb.AppendLine("$builder = New-Object Aspose.Words.DocumentBuilder $doc");
            sb.AppendLine();
            sb.AppendLine("$imageFiles = Get-ChildItem -Path $imagesFolder -File");
            sb.AppendLine("foreach ($img in $imageFiles) {");
            sb.AppendLine("    $builder.Writeln(\"Re‑embedding image: $($img.Name)\")");
            sb.AppendLine("    $builder.InsertImage($img.FullName)");
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("$doc.Save($outputPath)");
            sb.AppendLine("Write-Host \"Document saved to $outputPath\"");

            File.WriteAllText(scriptPath, sb.ToString());
        }

        // Helper to escape backslashes for PowerShell string literals.
        private static string EscapeForPowerShell(string path)
        {
            return path.Replace("\\", "\\\\");
        }
    }
}
