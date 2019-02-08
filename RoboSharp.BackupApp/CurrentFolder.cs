using System;

namespace RoboSharp.BackupApp {
  public class CurrentFolder {
    private string _name;
    private string _fileName;

    public string Name { get => _name ?? ""; set => _name = value; }  // Name of folder
    public long FilesTotal { get; set; } // Number of files in folder record
    public long FilesSkipped { get; set; } // Current files count
    public long FilesCopied { get; set; } // Current files count
    public long BytesSkipped { get; set; }
    public long BytesCopied { get; set; }
    public string FileName { get => _fileName ?? ""; set => _fileName = value; } // Current file name
    public long FileSize { get; set; } // Current file Size
    public long FilePortion { get; set; } // Current file Portion
    public double FilePC { get; set; } // Current file PC
    public bool ShowFile { get; set; } // File is to be coppied 

    // used in tf BytesCopied FilesCopied FilesSkipped BytesSkipped
  }
  public class Totals {
    public long FilesSkipped { get; set; } // Current files count
    public long FilesCopied { get; set; } // Current files count
    public long BytesSkipped { get; set; }
    public long BytesCopied { get; set; }
    public long ExtraFileCount { get; set; }
    public long ExtraByteCount { get; set; }
    public long FolderCount { get; set; }
    public long TotalFolders { get; set; }
    //public long ShowFolders { get; set; }
    public long ErrorCount { get; set; }
    //public long  { get; set; }
}



class Metric {
    public string Type { get; set; }
    public string Total { get; set; }
    public string Copied { get; set; }
    public string Skipped { get; set; }
    public string Mismatch { get; set; }
    public string FAILED { get; set; }
    public string Extras { get; set; }
    public override string ToString() {
      return $"{Type} T{Total} C{Copied} S{Skipped} M{Mismatch} F{FAILED} E{Extras}";
    }
  }

  public class FileError {
    public string Error { get; set; }
  }
}
