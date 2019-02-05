using System;

namespace RoboSharp.BackupApp {
  public class Folder {
    public string Name { get; set; }  // Name of folder
    public long FilesTotal { get; set; } // Number of files in folder record
    public long FilesSkipped { get; set; } // Current files count
    public long FilesCopied { get; set; } // Current files count
    public long BytesSkipped { get; set; }
    public long BytesCopied { get; set; } 
    public string FileName { get; set; } // Current file name
    public long FileSize { get; set; } // Current file Size
    public long FilePortion { get; set; } // Current file Size
    public bool ShowFile { get; set; } // File is to be coppied 
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
