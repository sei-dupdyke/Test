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
