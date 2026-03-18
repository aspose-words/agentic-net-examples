using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Rendering;

namespace OfficeMathMarginValidation
{
    class Program
    {
        static void Main()
        {
            // Load the source document.
            // (Assuming the document is located at the given path.)
            Document doc = new Document("Input.docx");

            // Bulk change justification for all OfficeMath objects.
            // Example: set all equations to be centered as a group.
            foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true))
            {
                // Ensure the display type is set before changing justification.
                officeMath.DisplayType = OfficeMathDisplayType.Display;
                officeMath.Justification = OfficeMathJustification.CenterGroup;
            }

            // Recalculate the layout so that size and position information is up‑to‑date.
            doc.UpdatePageLayout();

            // Retrieve page dimensions and margins from the first section.
            // (All sections are assumed to have the same layout for this validation.)
            PageSetup pageSetup = doc.FirstSection.PageSetup;
            double pageWidth = pageSetup.PageWidth;          // Total page width in points.
            double leftMargin = pageSetup.LeftMargin;        // Left margin in points.
            double rightMargin = pageSetup.RightMargin;      // Right margin in points.
            double usableWidth = pageWidth - leftMargin - rightMargin; // Width available for content.

            // Validate each OfficeMath equation does not exceed the usable page width.
            List<OfficeMath> offendingEquations = new List<OfficeMath>();
            foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true))
            {
                // Render the equation to obtain its size in points.
                OfficeMathRenderer renderer = new OfficeMathRenderer(officeMath);
                double equationWidth = renderer.SizeInPoints.Width;

                // If the equation width is greater than the usable width, record it.
                if (equationWidth > usableWidth)
                    offendingEquations.Add(officeMath);
            }

            // Report validation result.
            if (offendingEquations.Count == 0)
            {
                Console.WriteLine("All OfficeMath equations fit within the page margins.");
            }
            else
            {
                Console.WriteLine($"Found {offendingEquations.Count} equation(s) that exceed page margins:");
                foreach (OfficeMath om in offendingEquations)
                {
                    // Output the equation text for debugging purposes.
                    Console.WriteLine($"- Equation text: \"{om.GetText().Trim()}\"");
                }
            }

            // Save the modified document (optional).
            doc.Save("Output.docx");
        }
    }
}
