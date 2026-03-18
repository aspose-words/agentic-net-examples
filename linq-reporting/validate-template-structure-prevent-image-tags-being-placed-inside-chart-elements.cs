using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace TemplateValidation
{
    /// <summary>
    /// Provides methods to validate that no image tags are placed inside chart elements of a Word template.
    /// </summary>
    public static class ChartImageValidator
    {
        /// <summary>
        /// Loads a Word document from the specified path and validates that chart shapes do not contain image shapes
        /// or image template tags (e.g., "<<image>>") within their titles.
        /// </summary>
        /// <param name="templatePath">Full path to the template document.</param>
        public static void ValidateNoImagesInCharts(string templatePath)
        {
            // Load the document.
            Document doc = new Document(templatePath);

            // Get all shapes in the document (including those inside groups).
            NodeCollection allShapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in allShapes.OfType<Shape>())
            {
                // Process only shapes that contain a chart.
                if (!shape.HasChart)
                    continue;

                // 1. Ensure the chart shape itself is not an image shape.
                if (shape.IsImage)
                    throw new InvalidOperationException($"Shape '{shape.Name}' is marked as a chart but also reports IsImage = true.");

                // 2. Check the chart title for image template tags.
                //    The title is stored as plain text; if it contains a reporting engine tag that inserts an image,
                //    we consider it invalid.
                Chart chart = shape.Chart;
                if (chart?.Title != null && !string.IsNullOrEmpty(chart.Title.Text))
                {
                    string titleText = chart.Title.Text;
                    if (titleText.Contains("<<image") || titleText.Contains("{{image"))
                        throw new InvalidOperationException($"Chart '{titleText}' contains an image tag in its title.");
                }

                // 3. Search for any descendant image shapes inside the chart shape.
                //    Although a chart shape normally does not have child shapes, a group shape could be used.
                NodeCollection descendantShapes = shape.GetChildNodes(NodeType.Shape, true);
                foreach (Shape innerShape in descendantShapes.OfType<Shape>())
                {
                    if (innerShape.IsImage)
                        throw new InvalidOperationException($"Image shape '{innerShape.Name}' found inside chart shape '{shape.Name}'.");
                }
            }

            Console.WriteLine("Template validation succeeded: no image tags are placed inside chart elements.");
        }

        // Example usage.
        public static void Main()
        {
            // Path to the template document to be validated.
            string templateFile = Path.Combine(Environment.CurrentDirectory, "Template.docx");

            // Perform validation.
            ValidateNoImagesInCharts(templateFile);
        }
    }
}
