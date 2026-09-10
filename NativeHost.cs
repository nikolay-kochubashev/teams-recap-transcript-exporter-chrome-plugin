using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

public static class TeamsRecapNativeHost
{
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

    public static int Main(string[] args)
    {
        Console.InputEncoding = Encoding.UTF8;
        Console.OutputEncoding = Encoding.UTF8;

        try
        {
            Stream input = Console.OpenStandardInput();
            Stream output = Console.OpenStandardOutput();
            while (true)
            {
                string raw = ReadMessage(input);
                if (raw == null) break;
                Dictionary<string, object> request;
                try
                {
                    request = Json.Deserialize<Dictionary<string, object>>(raw);
                    WriteMessage(output, Handle(request));
                }
                catch (Exception ex)
                {
                    WriteMessage(output, new Dictionary<string, object> {
                        { "ok", false }, { "error", ex.Message }
                    });
                }
            }
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }

    private static Dictionary<string, object> Handle(Dictionary<string, object> request)
    {
        string action = GetString(request, "action");
        if (action == "ping")
        {
            return new Dictionary<string, object> {
                { "ok", true }, { "documents", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) }
            };
        }

        if (action == "saveText")
        {
            string fileName = SanitizeFileName(GetString(request, "fileName"));
            string text = GetString(request, "text");
            if (String.IsNullOrWhiteSpace(fileName)) fileName = "teams-transcript.txt";
            if (!fileName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) fileName += ".txt";

            string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (String.IsNullOrWhiteSpace(docs)) throw new InvalidOperationException("Windows Documents folder is unavailable.");
            Directory.CreateDirectory(docs);
            string path = Path.Combine(docs, fileName);
            File.WriteAllText(path, text ?? String.Empty, new UTF8Encoding(false));

            return new Dictionary<string, object> {
                { "ok", true }, { "path", path }, { "directory", docs }
            };
        }

        if (action == "showInFolder")
        {
            string path = GetString(request, "path");
            if (String.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new FileNotFoundException("Saved file was not found.", path);

            Process.Start(new ProcessStartInfo {
                FileName = "explorer.exe",
                Arguments = "/select,\"" + path.Replace("\"", "") + "\"",
                UseShellExecute = true
            });
            return new Dictionary<string, object> { { "ok", true } };
        }

        if (action == "openFile")
        {
            string path = GetString(request, "path");
            if (String.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new FileNotFoundException("Saved file was not found.", path);
            Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
            return new Dictionary<string, object> { { "ok", true } };
        }

        throw new InvalidOperationException("Unknown action: " + action);
    }

    private static string GetString(Dictionary<string, object> request, string key)
    {
        object value;
        if (request != null && request.TryGetValue(key, out value) && value != null) return Convert.ToString(value);
        return String.Empty;
    }

    private static string SanitizeFileName(string fileName)
    {
        string value = fileName ?? String.Empty;
        foreach (char c in Path.GetInvalidFileNameChars()) value = value.Replace(c, '-');
        value = value.Trim().TrimEnd('.', ' ');
        if (value.Length > 180) value = value.Substring(0, 180).TrimEnd('.', ' ');
        return value;
    }

    private static string ReadMessage(Stream input)
    {
        byte[] lenBytes = new byte[4];
        int first = input.ReadByte();
        if (first < 0) return null;
        lenBytes[0] = (byte)first;
        ReadExact(input, lenBytes, 1, 3);
        int length = BitConverter.ToInt32(lenBytes, 0);
        if (length < 0 || length > 64 * 1024 * 1024) throw new InvalidDataException("Invalid native message length.");
        byte[] data = new byte[length];
        ReadExact(input, data, 0, length);
        return Encoding.UTF8.GetString(data);
    }

    private static void WriteMessage(Stream output, object response)
    {
        byte[] data = Encoding.UTF8.GetBytes(Json.Serialize(response));
        byte[] len = BitConverter.GetBytes(data.Length);
        output.Write(len, 0, len.Length);
        output.Write(data, 0, data.Length);
        output.Flush();
    }

    private static void ReadExact(Stream stream, byte[] buffer, int offset, int count)
    {
        while (count > 0)
        {
            int read = stream.Read(buffer, offset, count);
            if (read <= 0) throw new EndOfStreamException();
            offset += read;
            count -= read;
        }
    }
}
