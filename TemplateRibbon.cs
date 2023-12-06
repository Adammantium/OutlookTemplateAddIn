using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;

[ComVisible(true)]
public partial class TemplateRibbon : IRibbonExtensibility {
    private IRibbonUI ribbon;

    public string GetCustomUI(string ribbonID) {
        if(ribbonID == "Microsoft.Outlook.Explorer") {
            return GenerateXml();
        }
        return null;
    }

    public void OnLoad(IRibbonUI ribbonUI) {
        this.ribbon = ribbonUI;
    }

    private Dictionary<int, FileInfo> files = new Dictionary<int, FileInfo>();
    private Dictionary<int, FileInfo> customIcons = new Dictionary<int, FileInfo>();

    private int GetId(string controlId) {
        try {
            return int.Parse(controlId.Substring("CustomButton".Length));
        }catch(Exception) {
            return -1;
        }
    }

    public Bitmap GetCustomImage(IRibbonControl control) {
        int fileId = GetId(control.Id);
        if(fileId == -1) return null;

        Bitmap bmp = new Bitmap(customIcons[fileId].FullName);
        return bmp;
    }

    public void OnButtonClick(IRibbonControl control) {
        int fileId = GetId(control.Id);
        if(fileId == -1) return;

        FileInfo file = files[fileId];
        if(file.Exists) {
            OpenWithDefaultProgram(file.FullName);
        }
    }

    public static void OpenWithDefaultProgram(string path) {
        Process fileopener = new Process();

        fileopener.StartInfo.FileName = "explorer";
        fileopener.StartInfo.Arguments = "\"" + path + "\"";
        fileopener.Start();
    }

    private string GenerateXml() {
        string header = @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""RibbonLoad"">
                        <ribbon>
                        <tabs>";

        string footer = @"</tabs>
                        </ribbon>
                        </customUI>";

        files.Clear();

        int fileCounter = 0;
        int categoryCounter = 0;
        string xml = "";
        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookTemplates");
        foreach(string tabPath in Directory.GetDirectories(dir)) {
            bool valid = false;
            string tabName = Path.GetFileName(tabPath);
            string tabXml = "<tab id=\"" + tabName.Replace(" ", "") + "\" label=\"" + tabName + "\">";
            foreach(string categoryPath in Directory.GetDirectories(tabPath)) {
                string categoryName = Path.GetFileName(categoryPath);
                tabXml += "<group id=\"Category" + ++categoryCounter + "\" label=\"" + categoryName + "\">";
                foreach(FileInfo file in GetFilesByExtensions(categoryPath, ".msg", ".oft", ".txt", ".html", ".lnk", ".url")) {
                    fileCounter++;

                    bool hasCustomIcon = false;
                    if(File.Exists(file.FullName + ".png")) {
                        customIcons.Add(fileCounter, new FileInfo(file.FullName + ".png"));
                        hasCustomIcon = true;
                    }

                    tabXml += $"<button id=\"CustomButton{ fileCounter }\" " +
                              $"label=\"{ Path.GetFileNameWithoutExtension(file.Name) }\" " +
                              (hasCustomIcon ? "getImage=\"GetCustomImage\"" : $"imageMso=\"{GetIconForExtension(file.Extension)}\"") + " " +
                              $"size=\"large\" onAction=\"OnButtonClick\" " +
                              $"/>";
                    files.Add(fileCounter, file);
                    valid = true;
                }
                tabXml += "</group>";
            }
            tabXml += "</tab>";

            if(valid) xml += tabXml;
        }

        xml = header + xml + footer;

        return xml;
    }

    private string GetIconForExtension(string extension) {
        switch (extension) {
            case ".msg":
            case ".oft":
                return "NewMailMessage";

            case ".txt":
                return "TextFromFileInsert";

            case ".html":
            case ".url":
                return "RmsInvokeBrowser";

            default: return "FileNew";
        }
    }

    private IEnumerable<FileInfo> GetFilesByExtensions(string dir, params string[] extensions) {
        if(extensions == null)
            throw new ArgumentNullException("extensions");
        DirectoryInfo directoryInfo = new DirectoryInfo(dir);

        FileInfo[] files = directoryInfo.GetFiles();
        return files.Where(f => extensions.Contains(f.Extension));
    }
}
