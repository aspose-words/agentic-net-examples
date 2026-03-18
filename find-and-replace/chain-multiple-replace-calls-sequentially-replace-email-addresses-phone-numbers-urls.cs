using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // 1. Replace all e‑mail addresses with a placeholder.
        // Pattern matches typical e‑mail formats, case‑insensitive.
        doc.Range.Replace(
            new Regex(@"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", RegexOptions.IgnoreCase),
            "[email]");

        // 2. Replace phone numbers with a placeholder.
        // Simple pattern that covers formats like 123-456-7890, (123) 456‑7890, +1 123 456 7890, etc.
        doc.Range.Replace(
            new Regex(@"\b(?:\+?\d{1,3}[-.\s]?)?(?:\(?\d{3}\)?[-.\s]?)?\d{3}[-.\s]?\d{4}\b"),
            "[phone]");

        // 3. Replace URLs with a placeholder.
        // Matches http, https, or www prefixes followed by non‑space characters.
        doc.Range.Replace(
            new Regex(@"\b(?:https?://|www\.)\S+\b", RegexOptions.IgnoreCase),
            "[url]");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
