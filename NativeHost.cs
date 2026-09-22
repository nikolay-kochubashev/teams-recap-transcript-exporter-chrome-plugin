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

                try
                {
                    Dictionary<string, object> request = Json.Deserialize<Dictionary<string, object>>(raw);
                    WriteMessage(output, Handle(request));
                }
                catch (Exception ex)
                {
                    WriteMessage(output, Error(ex));
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
            return Ok(new Dictionary<string, object> {
                { "documents", DocumentsPath() },
                { "batchRoot", BatchRootPath() },
                { "version", "2.0.0" }
            });
        }

        if (action == "saveText")
        {
            string fileName = EnsureTxt(SanitizeFileName(GetString(request, "fileName")));
            string path = Path.Combine(DocumentsPath(), fileName);
            File.WriteAllText(path, GetString(request, "text"), new UTF8Encoding(false));

            return Ok(new Dictionary<string, object> {
                { "path", path },
                { "directory", DocumentsPath() }
            });
        }

        if (action == "createBatchFolder")
        {
            string root = BatchRootPath();
            Directory.CreateDirectory(root);

            string baseName = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            string path = Path.Combine(root, baseName);
            int suffix = 2;

            while (Directory.Exists(path))
            {
                path = Path.Combine(root, baseName + "-" + suffix.ToString("00"));
                suffix++;
            }

            Directory.CreateDirectory(path);

            return Ok(new Dictionary<string, object> {
                { "path", path },
                { "root", root }
            });
        }

        if (action == "saveTextInFolder")
        {
            string folderPath = ValidateBatchFolder(GetString(request, "folderPath"));
            string fileName = EnsureTxt(SanitizeFileName(GetString(request, "fileName")));
            string path = Path.Combine(folderPath, fileName);

            File.WriteAllText(path, GetString(request, "text"), new UTF8Encoding(false));

            return Ok(new Dictionary<string, object> {
                { "path", path },
                { "directory", folderPath }
            });
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

            return Ok(null);
        }

        if (action == "openDirectory")
        {
            string path = ValidateAllowedDirectory(GetString(request, "path"));

            if (!Directory.Exists(path))
                throw new DirectoryNotFoundException("Directory was not found: " + path);

            Process.Start(new ProcessStartInfo {
                FileName = "explorer.exe",
                Arguments = "\"" + path.Replace("\"", "") + "\"",
                UseShellExecute = true
            });

            return Ok(new Dictionary<string, object> { { "path", path } });
        }

        if (action == "openFile")
        {
            string path = GetString(request, "path");

            if (String.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new FileNotFoundException("Saved file was not found.", path);

            Process.Start(new ProcessStartInfo {
                FileName = path,
                UseShellExecute = true
            });

            return Ok(null);
        }

        throw new InvalidOperationException("Unknown action: " + action);
    }

    private static string DocumentsPath()
    {
        string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        if (String.IsNullOrWhiteSpace(docs))
            throw new InvalidOperationException("Windows Documents folder is unavailable.");

        Directory.CreateDirectory(docs);
        return Path.GetFullPath(docs);
    }

    private static string BatchRootPath()
    {
        return Path.Combine(DocumentsPath(), "Teams Transcripts");
    }

    private static string ValidateBatchFolder(string path)
    {
        if (String.IsNullOrWhiteSpace(path))
            throw new InvalidOperationException("Batch folder path is empty.");

        string full = Path.GetFullPath(path);
        string root = Path.GetFullPath(BatchRootPath()).TrimEnd(Path.DirectorySeparatorChar)
            + Path.DirectorySeparatorChar;

        if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "Batch folder must be inside Windows Documents\\Teams Transcripts."
            );

        Directory.CreateDirectory(full);
        return full;
    }

    private static string ValidateAllowedDirectory(string path)
    {
        if (String.IsNullOrWhiteSpace(path))
            throw new InvalidOperationException("Directory path is empty.");

        string full = Path.GetFullPath(path);
        string documents = Path.GetFullPath(DocumentsPath());
        string documentsPrefix = documents.TrimEnd(Path.DirectorySeparatorChar)
            + Path.DirectorySeparatorChar;

        bool isDocuments = String.Equals(
            full.TrimEnd(Path.DirectorySeparatorChar),
            documents.TrimEnd(Path.DirectorySeparatorChar),
            StringComparison.OrdinalIgnoreCase
        );

        if (!isDocuments && !full.StartsWith(documentsPrefix, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "Only directories inside Windows Documents are allowed."
            );

        return full;
    }

    private static Dictionary<string, object> Ok(Dictionary<string, object> values)
    {
        Dictionary<string, object> result = values ?? new Dictionary<string, object>();
        result["ok"] = true;
        return result;
    }

    private static Dictionary<string, object> Error(Exception ex)
    {
        return new Dictionary<string, object> {
            { "ok", false },
            { "error", ex.Message }
        };
    }

    private static string GetString(Dictionary<string, object> request, string key)
    {
        object value;

        if (request != null && request.TryGetValue(key, out value) && value != null)
            return Convert.ToString(value);

        return String.Empty;
    }

    private static string EnsureTxt(string fileName)
    {
        if (String.IsNullOrWhiteSpace(fileName))
            fileName = "teams-transcript.txt";

        if (!fileName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
            fileName += ".txt";

        return fileName;
    }

    private static string SanitizeFileName(string fileName)
    {
        string value = fileName ?? String.Empty;

        foreach (char c in Path.GetInvalidFileNameChars())
            value = value.Replace(c, '-');

        value = value.Trim().TrimEnd('.', ' ');

        if (value.Length > 180)
            value = value.Substring(0, 180).TrimEnd('.', ' ');

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

        if (length < 0 || length > 64 * 1024 * 1024)
            throw new InvalidDataException("Invalid native message length.");

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

            if (read <= 0)
                throw new EndOfStreamException();

            offset += read;
            count -= read;
        }
    }
}
