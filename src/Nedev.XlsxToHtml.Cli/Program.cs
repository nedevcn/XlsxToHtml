using System;
using System.IO;
using Nedev.XlsxToHtml;

if (args.Length < 1 || args.Length > 2)
{
    Console.WriteLine("Usage: dotnet run --project src/Nedev.XlsxToHtml.Cli -- <input.xlsx> [output.html]");
    Console.WriteLine("       dotnet run --project src/Nedev.XlsxToHtml.Cli -- --dump-strings <input.xlsx>");
    return;
}

string input = args[0];
string output = args.Length == 2 ? args[1] : null;

// handle debug option
if (input == "--dump-strings")
{
    if (args.Length < 2)
    {
        Console.Error.WriteLine("Missing filename for --dump-strings");
        return;
    }
    input = args[1];
    var reader2 = new XlsxReader();
    var workbook2 = reader2.Read(input);
    // read shared strings directly from the archive
    using var archive = System.IO.Compression.ZipFile.OpenRead(input);
    var list = new System.Collections.Generic.List<string>();
    var entry = archive.GetEntry("xl/sharedStrings.xml");
    if (entry != null)
    {
        using var stream = entry.Open();
        using var reader3 = System.Xml.XmlReader.Create(stream);
        while (reader3.Read())
        {
            if (reader3.NodeType == System.Xml.XmlNodeType.Element && reader3.Name == "si")
            {
                var sb = new System.Text.StringBuilder();
                while (reader3.Read())
                {
                    if (reader3.NodeType == System.Xml.XmlNodeType.Element && reader3.Name == "t")
                    {
                        sb.Append(reader3.ReadElementContentAsString());
                    }
                    else if (reader3.NodeType == System.Xml.XmlNodeType.EndElement && reader3.Name == "si")
                    {
                        break;
                    }
                }
                list.Add(sb.ToString());
            }
        }
    }
    Console.WriteLine($"Shared strings count: {list.Count}");
    for (int i = 0; i < Math.Min(list.Count, 20); i++)
        Console.WriteLine($"{i}: {list[i]}");
    return;
}

if (!File.Exists(input))
{
    Console.Error.WriteLine($"Input file does not exist: {input}");
    Environment.Exit(1);
}

try
{
    var sw = System.Diagnostics.Stopwatch.StartNew();
    var reader = new XlsxReader();
    var workbook = reader.Read(input);
    var writer = new HtmlWriter();
    if (string.IsNullOrEmpty(output))
    {
        // write to stdout
        writer.Write(workbook, Console.Out);
    }
    else
    {
        using var fs = new StreamWriter(output, false, System.Text.Encoding.UTF8);
        writer.Write(workbook, fs);
    }
    sw.Stop();
    Console.Error.WriteLine($"Converted in {sw.ElapsedMilliseconds} ms");
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    Environment.Exit(2);
}
