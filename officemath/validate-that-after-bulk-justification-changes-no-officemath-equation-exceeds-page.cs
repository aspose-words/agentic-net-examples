using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Math;

namespace OfficeMathMarginValidation
{
    class Program
    {
        static void Main()
        {
            // Load the source document.
            Document doc = new Document("Input.docx");

            // Bulk change justification for all OfficeMath objects.
            foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true).Cast<OfficeMath>())
            {
                officeMath.DisplayType = OfficeMathDisplayType.Display;
                officeMath.Justification = OfficeMathJustification.CenterGroup;
            }

            // Retrieve page dimensions and margins from the first section.
            PageSetup pageSetup = doc.FirstSection.PageSetup;
            double pageWidth = pageSetup.PageWidth;          // Total page width in points.
            double leftMargin = pageSetup.LeftMargin;        // Left margin in points.
            double rightMargin = pageSetup.RightMargin;      // Right margin in points.
            double usableWidth = pageWidth - leftMargin - rightMargin; // Width available for content.

            // Validate each OfficeMath equation does not exceed the usable page width.
            List<OfficeMath> offendingEquations = new List<OfficeMath>();
            foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true).Cast<OfficeMath>())
            {
                // Approximate the equation width using character count.
                // This is a placeholder for actual rendering size calculation.
                double approximateCharWidth = 5.0; // Approximate width of a character in points.
                double equationWidth = officeMath.GetText().Length * approximateCharWidth;

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
                    Console.WriteLine($"- Equation text: \"{om.GetText().Trim()}\"");
                }
            }

            // Save the modified document (optional).
            doc.Save("Output.docx");
        }
    }
}
