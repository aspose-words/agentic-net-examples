using System;
using Aspose.Words.Replacing;
using Aspose.Words.Shaping;
using Aspose.Words;
using Aspose.Words.Properties;
using Aspose.Words.Fields;

namespace Aspose.Words.Extensions
{
    /// <summary>
    /// Provides extension methods for enum types used in Aspose.Words.
    /// </summary>
    public static class EnumExtensions
    {
        /// <summary>
        /// Returns a user‑friendly string representation of the enum value.
        /// For known Aspose.Words enums the method returns a more readable text,
        /// otherwise it falls back to the default <c>ToString()</c> implementation.
        /// </summary>
        /// <param name="value">The enum value.</param>
        /// <returns>A friendly string.</returns>
        public static string ToFriendlyString(this Enum value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));

            // Handle specific Aspose.Words enums with custom friendly names.
            switch (value)
            {
                case ReplacementFormat rf:
                    return rf switch
                    {
                        ReplacementFormat.Text => "Plain Text",
                        ReplacementFormat.Markdown => "Markdown",
                        ReplacementFormat.Html => "HTML",
                        _ => rf.ToString()
                    };

                case Direction dir:
                    // Direction.Default and Direction.LTR share the same underlying value (0).
                    // Combine them in a single pattern to avoid unreachable‑code warnings.
                    return dir switch
                    {
                        Direction.Default or Direction.LTR => "Left‑to‑Right",
                        Direction.RTL => "Right‑to‑Left",
                        Direction.TTB => "Top‑to‑Bottom",
                        Direction.BTT => "Bottom‑to‑Top",
                        _ => dir.ToString()
                    };

                case SaveFormat sf:
                    return sf switch
                    {
                        SaveFormat.Doc => "Word 97‑2003 Document (.doc)",
                        SaveFormat.Docx => "Word Document (.docx)",
                        SaveFormat.Pdf => "PDF",
                        SaveFormat.Html => "HTML",
                        SaveFormat.Text => "Plain Text",
                        SaveFormat.Markdown => "Markdown",
                        SaveFormat.Png => "PNG Image",
                        SaveFormat.Jpeg => "JPEG Image",
                        SaveFormat.Tiff => "TIFF Image",
                        _ => sf.ToString()
                    };

                case LoadFormat lf:
                    return lf switch
                    {
                        LoadFormat.Doc => "Word 97‑2003 Document (.doc)",
                        LoadFormat.Docx => "Word Document (.docx)",
                        LoadFormat.Html => "HTML",
                        LoadFormat.Text => "Plain Text",
                        LoadFormat.Markdown => "Markdown",
                        _ => lf.ToString()
                    };

                case PropertyType pt:
                    return pt switch
                    {
                        PropertyType.Boolean => "Boolean",
                        PropertyType.DateTime => "Date/Time",
                        PropertyType.Double => "Floating‑Point Number",
                        PropertyType.Number => "Integer Number",
                        PropertyType.String => "String",
                        PropertyType.StringArray => "String Array",
                        PropertyType.ObjectArray => "Object Array",
                        PropertyType.ByteArray => "Byte Array",
                        PropertyType.Other => "Other",
                        _ => pt.ToString()
                    };

                case GeneralFormat gf:
                    return gf switch
                    {
                        GeneralFormat.None => "None",
                        GeneralFormat.Upper => "Uppercase",
                        GeneralFormat.Lower => "Lowercase",
                        GeneralFormat.Caps => "Capitalize Each Word",
                        GeneralFormat.FirstCap => "Capitalize First Word",
                        GeneralFormat.UppercaseRoman => "Uppercase Roman Numerals",
                        GeneralFormat.LowercaseRoman => "Lowercase Roman Numerals",
                        GeneralFormat.Ordinal => "Ordinal (1st, 2nd, …)",
                        _ => gf.ToString()
                    };

                default:
                    // For any other enum, just use its name.
                    return value.ToString();
            }
        }
    }

    /// <summary>
    /// Simple console entry point used only to satisfy the compiler and to demonstrate the extension.
    /// </summary>
    class Program
    {
        static void Main()
        {
            Console.WriteLine(ReplacementFormat.Markdown.ToFriendlyString());
            Console.WriteLine(Direction.LTR.ToFriendlyString());
            Console.WriteLine(SaveFormat.Pdf.ToFriendlyString());
            Console.WriteLine(LoadFormat.Docx.ToFriendlyString());
            Console.WriteLine(PropertyType.DateTime.ToFriendlyString());
            Console.WriteLine(GeneralFormat.Caps.ToFriendlyString());
        }
    }
}
