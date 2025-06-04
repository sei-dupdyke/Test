using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
#if WINDOWS
using Word = Microsoft.Office.Interop.Word;
#endif

class Program
{
    /// <summary>
    /// Creates a sample letter using the available automation method for the
    /// current platform. Windows uses COM Interop, macOS uses AppleScript, and
    /// other platforms fall back to generating a text file.
    /// <summary>
    /// Detects the current operating system and generates a sample letter using a platform-specific method.
    /// </summary>
    static void Main(string[] args)
    {
        string[] lines = new[]
        {
            "Dear Editor,",
            "",
            "I am writing to express my views on the recent events.",
            "Thank you for considering my letter.",
            "",
            "Sincerely,",
            "A Concerned Reader"
        };

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            WriteLetterWindows(lines);
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            WriteLetterMac(lines);
        }
        else
        {
            WriteLetterOther(lines);
        }
    }

#if WINDOWS
    /// <summary>
    /// Creates a new Microsoft Word document and simulates typing the provided lines character by character, adding a paragraph break after each line.
    /// </summary>
    /// <param name="lines">The lines of text to be typed into the Word document.</param>
    static void WriteLetterWindows(string[] lines)
    {
        var wordApp = new Word.Application();
        wordApp.Visible = true;
        var document = wordApp.Documents.Add();

        foreach (var line in lines)
        {
            foreach (char c in line)
            {
                wordApp.Selection.TypeText(c.ToString());
                System.Threading.Thread.Sleep(50);
            }
            wordApp.Selection.TypeParagraph();
        }
    }
#endif

    /// <summary>
    /// Creates a new Microsoft Word document on macOS using AppleScript and sets its content to the provided letter lines.
    /// </summary>
    /// <param name="lines">The lines of text to include in the letter.</param>
    static void WriteLetterMac(string[] lines)
    {
        var scriptPath = Path.Combine(Path.GetTempPath(), "letter.scpt");
        string joined = string.Join("\\n", lines).Replace("\"", "\\\"");

        var script = string.Join('\n', new[]
        {
            "tell application \"Microsoft Word\"",
            "activate",
            "set newDoc to make new document",
            $"set content of text object of newDoc to \"{joined}\"",
            "end tell"
        });

        File.WriteAllText(scriptPath, script);
        Process.Start("osascript", scriptPath);
    }

    /// <summary>
    /// Writes the provided lines to a temporary text file and opens it with the default application on non-Windows, non-macOS platforms.
    /// </summary>
    /// <param name="lines">The lines of the letter to write to the file.</param>
    static void WriteLetterOther(string[] lines)
    {
        string path = Path.Combine(Path.GetTempPath(), "letter.txt");
        File.WriteAllLines(path, lines);

        var psi = new ProcessStartInfo
        {
            FileName = path,
            UseShellExecute = true
        };
        Process.Start(psi);
    }
}
