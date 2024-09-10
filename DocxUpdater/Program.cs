using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace DocxUpdater
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            const string marker = "-=XML=-";
            const int check = 100;

            var embed = args.Contains("--embed") || args.Contains("/embed");
            var path = typeof(Program).Assembly.Location;
            string xmlContents;

            var xmlPath = args.FirstOrDefault(arg => arg.EndsWith(".xml"));
            if (xmlPath == null)
            {
                var contents = File.ReadAllText(path, Encoding.UTF8);
                var index = contents.IndexOf(marker, StringComparison.Ordinal);
                if (index == -1)
                {
                    var ofd = new OpenFileDialog { Title = "Select XML file", Filter = "XML files|*.xml" };
                    if (ofd.ShowDialog() == true)
                        xmlPath = ofd.FileName;

                    if (string.IsNullOrEmpty(xmlPath))
                        return;

                    xmlContents = File.ReadAllText(xmlPath, Encoding.UTF8);
                }
                else
                    xmlContents = contents.Substring(index + marker.Length + 1); // NOTE: Skip version
            }
            else
                xmlContents = File.ReadAllText(xmlPath, Encoding.UTF8);

            if (embed)
            {
                var target = Path.ChangeExtension(path, "." + Path.GetFileNameWithoutExtension(xmlPath) + ".exe");
                File.Copy(path, target, true);
                // NOTE: v1 saves XML as clear text
                File.AppendAllText(target, marker + '\x01' + xmlContents, Encoding.UTF8);
                return;
            }

            var docxPath = args.FirstOrDefault(arg => arg.EndsWith(".docx"));
            if (docxPath == null)
            {
                var ofd = new OpenFileDialog { Title = "Select Microsoft Word document", Filter = "Microsoft Word documents|*.docx;*.dotx" };
                if (ofd.ShowDialog() == true)
                    docxPath = ofd.FileName;

                if (string.IsNullOrEmpty(docxPath))
                    return;
            }

            // NOTE: Replace XML contents
            var start = xmlContents.Substring(0, check);
            using (var package = Package.Open(docxPath, FileMode.Open, FileAccess.ReadWrite))
            {
                var parts = package.GetParts().Where(p => p.ContentType == "application/xml" && p.Uri.OriginalString.StartsWith("/customXml")).ToArray();
                foreach (var xmlPart in parts)
                {
                    using (var stream = xmlPart.GetStream(FileMode.Open, FileAccess.ReadWrite))
                    using (var reader = new StreamReader(stream))
                    {
                        var buffer = new char[check];
                        if (reader.Read(buffer, 0, check) == check && new string(buffer) == start)
                        {
                            // NOTE: Found correct part
                            stream.Position = 0;
                            stream.SetLength(0);

                            using (var writer = new StreamWriter(stream))
                            {
                                writer.Write(xmlContents);
                                writer.Flush();
                            }

                            break;
                        }
                    }
                }

                package.Flush();
                package.Close();
            }
        }
    }
}