using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample chart image using Aspose.Drawing.
        // -----------------------------------------------------------------
        const int chartWidth = 400;
        const int chartHeight = 300;
        const string chartImagePath = "sample-chart.png";

        using (Bitmap bitmap = new Bitmap(chartWidth, chartHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(Color.White);

                // Draw simple column chart bars.
                int barCount = 5;
                int barWidth = chartWidth / (barCount * 2);
                Random rnd = new Random();

                for (int i = 0; i < barCount; i++)
                {
                    int barHeight = rnd.Next(50, chartHeight - 50);
                    int x = (i * 2 + 1) * barWidth;
                    int y = chartHeight - barHeight;

                    using (SolidBrush brush = new SolidBrush(Color.FromArgb(100 + i * 30, 150, 200)))
                    {
                        graphics.FillRectangle(brush, x, y, barWidth, barHeight);
                    }

                    using (Pen pen = new Pen(Color.Black, 2))
                    {
                        graphics.DrawRectangle(pen, x, y, barWidth, barHeight);
                    }
                }
            }

            // Save the chart image to a local file.
            bitmap.Save(chartImagePath);
        }

        // ---------------------------------------------------------------
        // 2. Create a DOCX document and insert the generated chart image.
        // ---------------------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the chart image.
        builder.InsertImage(chartImagePath);
        doc.Save(docPath);

        // ---------------------------------------------------------------
        // 3. Load the document and extract all images (including the chart).
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extractedPath = $"extracted-{imageIndex}.png";

                // Save the image data to a file.
                shape.ImageData.Save(extractedPath);
                imageIndex++;
            }
        }

        // ---------------------------------------------------------------
        // 4. Validate that at least one image was extracted.
        // ---------------------------------------------------------------
        if (imageIndex == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to PNG files.");
    }
}
