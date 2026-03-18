using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Math;

namespace OfficeMathPdfValidation
{
    public class Validator
    {
        /// <summary>
        /// Validates that the PDF exported from a DOCX retains the same OfficeMath layout.
        /// The validation checks that the number of OfficeMath objects is unchanged
        /// and that the rendering options that affect positioning are enabled.
        /// </summary>
        /// <param name="sourceDocxPath">Path to the source DOCX file containing OfficeMath.</param>
        /// <param name="outputPdfPath">Path where the PDF will be saved.</param>
        public static void ValidateOfficeMathPdfPositioning(string sourceDocxPath, string outputPdfPath)
        {
            // Load the source DOCX document.
            Document srcDoc = new Document(sourceDocxPath);

            // Count OfficeMath objects in the source document.
            int srcMathCount = srcDoc.GetChildNodes(NodeType.OfficeMath, true).Count;

            // Ensure that each top‑level OfficeMath is displayed as a block (Display) and left‑justified.
            // This makes the positioning explicit and reproducible in the PDF.
            foreach (OfficeMath math in srcDoc.GetChildNodes(NodeType.OfficeMath, true))
            {
                // Only top‑level OfficeMath can have its DisplayType changed.
                if (math.MathObjectType == MathObjectType.OMathPara)
                {
                    math.DisplayType = OfficeMathDisplayType.Display;
                    math.Justification = OfficeMathJustification.Left;
                }
            }

            // Configure PDF save options to preserve additional text positioning operators.
            // This option helps keep the exact layout of equations.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AdditionalTextPositioning = true
            };

            // Save the document as PDF using the configured options.
            srcDoc.Save(outputPdfPath, pdfOptions);

            // Load the generated PDF back into Aspose.Words.
            Document pdfDoc = new Document(outputPdfPath);

            // Count OfficeMath objects in the PDF document.
            int pdfMathCount = pdfDoc.GetChildNodes(NodeType.OfficeMath, true).Count;

            // Simple validation: the count of OfficeMath objects should be identical.
            if (srcMathCount != pdfMathCount)
                throw new InvalidOperationException(
                    $"OfficeMath count mismatch. Source: {srcMathCount}, PDF: {pdfMathCount}");

            // Additional validation can be performed by comparing the rendered size of each equation.
            // Here we compare the bounding box size (in points) of each top‑level OfficeMath.
            for (int i = 0; i < srcMathCount; i++)
            {
                OfficeMath srcMath = (OfficeMath)srcDoc.GetChild(NodeType.OfficeMath, i, true);
                OfficeMath pdfMath = (OfficeMath)pdfDoc.GetChild(NodeType.OfficeMath, i, true);

                // Only compare display equations (top‑level).
                if (srcMath.MathObjectType != MathObjectType.OMathPara ||
                    pdfMath.MathObjectType != MathObjectType.OMathPara)
                    continue;

                var srcRenderer = srcMath.GetMathRenderer();
                var pdfRenderer = pdfMath.GetMathRenderer();

                // Width and height should match within a small tolerance.
                const float tolerance = 0.5f; // points
                if (Math.Abs(srcRenderer.SizeInPoints.Width - pdfRenderer.SizeInPoints.Width) > tolerance ||
                    Math.Abs(srcRenderer.SizeInPoints.Height - pdfRenderer.SizeInPoints.Height) > tolerance)
                {
                    throw new InvalidOperationException(
                        $"OfficeMath size mismatch at index {i}. " +
                        $"Source (W:{srcRenderer.SizeInPoints.Width}, H:{srcRenderer.SizeInPoints.Height}) vs " +
                        $"PDF (W:{pdfRenderer.SizeInPoints.Width}, H:{pdfRenderer.SizeInPoints.Height})");
                }
            }

            // If no exception was thrown, the validation succeeded.
            Console.WriteLine("PDF validation succeeded: OfficeMath equations retain their positioning.");
        }

        // Example usage.
        public static void Main()
        {
            string docxPath = Path.Combine(Environment.CurrentDirectory, "Input", "OfficeMathSample.docx");
            string pdfPath = Path.Combine(Environment.CurrentDirectory, "Output", "OfficeMathSample.pdf");

            ValidateOfficeMathPdfPositioning(docxPath, pdfPath);
        }
    }
}
