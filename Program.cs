using System;
using System.Diagnostics;
using Microsoft.Win32;
using System.Net;
using System.IO;
using HtmlAgilityPack;
using Newtonsoft.Json;
using System.Xml;
using System.Collections.Generic;
using System.IO.Compression;

public class Program
{
    static string VersionTextFile = "Version.info";

    static void Main(string[] args)
    {
        try
        {
            if (NeedsUpdate())
            {
                Update();
            }
        }

        catch (Exception Ex)
        {
            Console.WriteLine(Ex.Message);
        }
    }

    static bool NeedsUpdate()
    {

        if (File.Exists(VersionTextFile) == false) return true;

        string Version = File.ReadAllText(VersionTextFile);

        if (Version != GetShortenedVersionNumber())
        {
            return true;
        }

        else
        {
            return false;
        }
    }

    static void Update()
    {
        foreach (var Process in Process.GetProcesses())
        {
            if (Process.ProcessName.ToLower().Contains("chromedriver"))
            {
                Process.Kill();
            }
        }

        if (File.Exists("chromedriver_win32.zip"))
        {
            File.Delete("chromedriver_win32.zip");
        }

        if (File.Exists("chromedriver.exe"))
        {
            File.Delete("chromedriver.exe");
        }

        string TargetVersion = GetShortenedVersionNumber();
        List<string> Downloads = GetDownloadList();

        string ExePath = AppDomain.CurrentDomain.BaseDirectory;

        string DownloadLink = "https://chromedriver.storage.googleapis.com/" + GetAppropriateVersion(TargetVersion, Downloads.ToArray());
        string DownloadPath = ExePath;

        string ReturnDownload = DownloadFile(DownloadLink, DownloadPath);
        string ReturnExtract = Extract(DownloadPath + "chromedriver_win32.zip", DownloadPath);

        File.WriteAllText(VersionTextFile, TargetVersion);
    }

    #region Update
    static string Extract(string ZipPath, string FilePath)
    {
        try
        {
            ZipFile.ExtractToDirectory(ZipPath, FilePath);
            return "";
        }

        catch (Exception Ex)
        {
            return Ex.Message;
        }
    }

    static string DownloadFile(string Link, string Path)
    {
        try
        {
            using (var client = new WebClient())
            {
                client.DownloadFile(Link, Path + "chromedriver_win32.zip");
            }

            return "";
        }

        catch (Exception Ex)
        {
            Console.WriteLine(Ex.ToString());
            return Ex.Message;
        }
    }

    static string GetAppropriateVersion(string TargetVersion, string[] Downloads)
    {
        foreach (var Download in Downloads)
        {
            if (Download.Split('.')[0] == TargetVersion)
            {
                return Download;
            }
        }

        return null;
    }

    public class ChromeDownloads
    {
        public string Key { get; set; }
    }

    public static string DownloadString(string address)
    {
        string text;
        using (var client = new WebClient())
        {
            text = client.DownloadString(address);
        }
        return text;
    }

    static string GetFullVersionNumber()
    {
        object path;
        path = Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", "", null);
        
        if (path != null)
        {
            return FileVersionInfo.GetVersionInfo(path.ToString()).FileVersion;
        }

        return null;
    }

    static string GetShortenedVersionNumber()
    {
        return GetFullVersionNumber().Split('.')[0];
    }

    static List<string> GetDownloadList(bool Verbose = false)
    {
        List<string> Links = new List<string>();

        #region Parse XML, Parse Json -> filteredString
        var XMLString = DownloadString("https://chromedriver.storage.googleapis.com/");
        XmlDocument doc = new XmlDocument();
        doc.LoadXml(XMLString);
        string jsonString = JsonConvert.SerializeXmlNode(doc);
        string filteredString = jsonString.Replace("\"?xml\":{\"@version\":\"1.0\",\"@encoding\":\"UTF-8\"},", "").Replace("{\"ListBucketResult\":{\"@xmlns\":\"http://doc.s3.amazonaws.com/2006-03-01\",\"Name\":\"chromedriver\",\"Prefix\":\"\",\"Marker\":\"\",\"IsTruncated\":\"false\",\"Contents\":", "");
        filteredString = filteredString.Substring(0, filteredString.Length - 2);
        #endregion

        ChromeDownloads[] chromeDownloads = JsonConvert.DeserializeObject<ChromeDownloads[]>(filteredString);

        foreach (var chromeDownload in chromeDownloads)
        {
            if (chromeDownload.Key.EndsWith("win32.zip"))
            {
                if (Verbose) Console.WriteLine(chromeDownload.Key);
                Links.Add(chromeDownload.Key);
            }
        }

        return Links;
    }
    #endregion

}