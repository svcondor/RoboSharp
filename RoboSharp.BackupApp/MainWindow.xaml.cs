﻿using System;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;
using Alphaleonis.Win32.Filesystem;
using System.Collections.Generic;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Media;
using System.Linq;

// Add NuGet AlphaFs https://alphafs.alphaleonis.com/
// Add Nuget Microsoft.Office.Interop.Excel
// Add Reference COM Microsoft Office xx.0 Object Library
// https://github.com/svcondor/RoboSharp
// OneNote Windows - Windows - Robocopy 2

namespace RoboSharp.BackupApp {

  public partial class MainWindow : Window {
    CurrentFolder cf = new CurrentFolder();  // Current folder info record
    Totals tf;  //  info record to total all folders
    RoboCommand copy;
    public ObservableCollection<FileError> Errors = new ObservableCollection<FileError>();
    string SourceLower;
    string DestinationLower;
    //List<string> Folders;
    //Regex regex = new Regex(@"[ ]{2,}", RegexOptions.None);
    bool totaling = false;
    bool backupIsRunning = false;
    bool prepareIsRunning = false;
    List<Metric> metrics;
    Stopwatch RunTimer = new Stopwatch();
    Stopwatch PauseTimer = new Stopwatch();
    string source;
    string exclude;
    OptionsControl optionsControl;
    double MbsMax;
    Excel.Application ExcelApp;

    public MainWindow() {
      InitializeComponent();
      optionsControl = OptionsControlI;
      this.Closing += MainWindow_Closing;
      this.ContentRendered += Window_ContentRendered;
      ErrorGrid.ItemsSource = Errors;
      RoboSharp.Debugger.Instance.DebugMessageEvent += DebugMessage;
    }

    void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
      if (copy != null) {
        if (backupIsRunning) {
          MessageBox.Show("Cancel the current backup before closing", "Backup");
          e.Cancel = true;
          return;
        }
        copy.Stop();
        copy.Dispose();
      }
    }

    void BackupButton_Click(object sender, RoutedEventArgs e) {
      totaling = false;
      MetricsGrid.ItemsSource = null;
      //OptionsGrid.IsEnabled = false;
      //ProgressTab.IsSelected = true;
      //ProgressGrid.IsEnabled = true;
      //CurrentFolder.Text = "Running Preliminary Scan...";
      CurrentFolder.Text = "";
      txtFiles.Text = "";
      txtMBytes.Text = "";
      txtFilesSkipped.Text = "";
      txtMBytesSkipped.Text = "";
      Errors.Clear();
      ErrorsTab.Header = "Errors";
      txtErrors.Text = "";
      txtErrors.Background = Brushes.White;
      txtFolders.Text = "";
      txtFilePc.Text = "";
      txtTotalTime.Text = "";
      copy = new RoboCommand();
      copy.OnFileProcessed += copy_OnFileProcessed;
      copy.OnCommandError += copy_OnCommandError;
      copy.OnError += copy_OnError;
      copy.OnCopyProgressChanged += copy_OnCopyProgressChanged;
      copy.OnCommandCompleted += copy_OnCommandCompleted;
      copy.CopyOptions.Source = Source.Text;
      copy.CopyOptions.Destination = Destination.Text;

      optionsControl.BuildRobocopyParameters(copy);
      //Debug.WriteLine("After BuildRobocopyParameters()");
      //Debug.WriteLine("After copy.CommandOptions");
      SourceLower = Source.Text.ToLower();
      DestinationLower = Destination.Text.ToLower();
      cf.FolderName = "";
      cf.FilesTotal = 0;
      cf.FilesCopied = 0;
      cf.ClearFileData();
      tf = new Totals();
      ScanThenCopy();
    }

    void DebugMessage(object sender, RoboSharp.Debugger.DebugMessageArgs e) {
      Console.WriteLine($"ROBO {e.Message}");
    }

    void copy_OnCommandError(object sender, ErrorEventArgs e) {
      Dispatcher.BeginInvoke((Action)(() => {
        MessageBox.Show(e.Error);
        //OptionsGrid.IsEnabled = true;
        //ProgressGrid.IsEnabled = false;
      }));
    }

    void copy_OnCopyProgressChanged(object sender, CopyProgressEventArgs e) {
      //if (e.CurrentFileProgress == 100) {
      //Debug.WriteLine($"Progress {e.CurrentFileProgress}");
      //}
      cf.FilePC = e.CurrentFileProgress;
    }

    void copy_OnError(object sender, ErrorEventArgs e) {
      Dispatcher.BeginInvoke((Action)(() => {
        try {
          Errors.Insert(0, new FileError { Error = e.Error });
          ErrorsTab.Header = $"Errors ({Errors.Count:#,##0})";
          txtErrors.Text = $"{(Errors.Count):#,##0}";
          txtErrors.Background = Brushes.Pink;
        }
        catch (Exception) {
          throw;
        }
      }));
    }

    void copy_OnFileProcessed(object sender, FileProcessedEventArgs e) {
      try {
        var file = e.ProcessedFile;
        //Debug.WriteLine($"C-{file.FileClass} T-{file.FileClassType} N-{file.Name} S-{file.Size}");
        if (cf.FileName != "") {
          if (cf.FilePortion != cf.FileSize) {
            tf.BytesCopied += (cf.FileSize - cf.FilePortion);
            cf.FilePortion = cf.FileSize;
          }
          //Debug.WriteLine($"FileProcessed2-\"{cf.FileName}\" {tf.BytesCopied} {cf.FileSize} {cf.FilePortion} {cf.FilePC}");
          cf.FileName = "";
        }

        if (file.FileClassType == FileClassType.NewDir) {
          if (file.FileClass == "New Dir") {
            if (file.Name.ToLower().StartsWith(DestinationLower)) {
            }
            else if (file.Name.ToLower().StartsWith(SourceLower)) {
              ++tf.FolderCount;
              cf.FolderName = file.Name;
              cf.FilesTotal = file.Size;
              cf.FilesCopied = 0;
              cf.ClearFileData();
            }
            else {
              throw new Exception($"Folder not in source or destination \"{file.Name}\"");
            }
          }
          else {
            throw new Exception($"Folder FileClass not \"New Dir\" \"{file.FileClass}\"");
          }
        }
        else if (file.FileClassType == FileClassType.File) {
          if (cf == null) {
            throw new Exception("cf == null");
          }
          if (file.FileClass == "New File"
            || file.FileClass == "Newer"
            || file.FileClass == "modified") {
            ++cf.FilesCopied;
            ++tf.FilesCopied;
            cf.FileSize = file.Size;
            if (cf.FileSize > 100 * 1024 * 1024) {
              cf.FileName = file.Name;
              cf.FilePortion = 0;
              cf.FilePC = 0.0;
            }
            else {
              cf.FileName = "";
              tf.BytesCopied += file.Size;
            }
            cf.ShowFile = true;
          }
          else if (file.FileClass == "same"
            || file.FileClass == "Older") {
            if (file.FileClass == "Older") {
            }
            cf.FileName = "";
            cf.FileSize = file.Size;
            ++cf.FilesCopied;
            ++tf.FilesSkipped;
            tf.BytesSkipped += file.Size;
            cf.ShowFile = false;
          }
          else if (file.FileClass == "*EXTRA File") {
            cf.FileName = "";
            ++tf.ExtraFileCount;
            tf.ExtraByteCount += file.Size;
            cf.ShowFile = false;
          }
          else {
            throw new Exception($"Unhandled FileClass \"{file.FileClass}\"");
          }
        }
        else if (file.FileClassType == FileClassType.SystemMessage) {
          Debug.WriteLine($"SYSTEM\t{file.Name}");

          if (file.Name.StartsWith("Total")) {
            totaling = true;
            metrics = new List<Metric>();
          }
          else if (totaling) {
            var splits = file.Name.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            string[] types = { "Dirs", "Files", "Bytes" };
            if (splits[1] == ":" && types.Contains(splits[0])) {
              var m1 = new Metric { Type = splits[0] };
              if (splits.Length > 2) m1.Total = splits[2];
              if (splits.Length > 3) m1.Copied = splits[3];
              if (splits.Length > 4) m1.Skipped = splits[4];
              if (splits.Length > 5) m1.Mismatch = splits[5];
              if (splits.Length > 6) m1.FAILED = splits[6];
              if (splits.Length > 7) m1.Extras = splits[7];
              metrics.Add(m1);
            }
          }
        }
        else {
          throw new Exception($"Unhandled FileClassType \"{file.FileClassType}\"");
        }
      }
      catch (Exception ex) {
        throw ex;
      }
    }

    void copy_OnCommandCompleted(object sender, RoboCommandCompletedEventArgs e) {
      Debug.WriteLine("CommandCompleted");
      if (cf.FileName != "") {
      }
      //Task.Run(() => ShowValues());
      Dispatcher.BeginInvoke((Action)(() => {
        backupIsRunning = false;
        RunTimer.Stop();
        PauseTimer.Stop();
        if (metrics?.Count > 0) {
          MetricsGrid.ItemsSource = metrics;
        }
        PauseButton.Content = "Pause";
        CancelButton.IsEnabled = false;
        PauseButton.IsEnabled = false;
        BackupButton.IsEnabled = true;

        //txtMBytes.Text = $"{(tf.BytesCopied / 1024 / 1024):#,##0}";
        //txtMBytesSkipped.Text = $"{(tf.BytesSkipped / 1024 / 1024):#,##0}";
        //txtFiles.Text = $"{tf.FilesCopied:#,##0}";
        //txtFilesSkipped.Text = $"{tf.FilesSkipped:#,##0}";
        //txtFolders.Text = $"{tf.FolderCount:#,##0}";
        //CurrentFolder.Text = cf.FolderName;
      }));

      Debug.WriteLine($"EVENT\tcopy_OnCommandCompleted");
      var ts = RunTimer.Elapsed;
      Debug.WriteLine($"End of Job {ts}");
      //Dispatcher.BeginInvoke((Action)(() => {
      //  //OptionsGrid.IsEnabled = true;
      //  //ProgressGrid.IsEnabled = false;
      //}));
    }

    void PauseButton_Click(object sender, RoutedEventArgs e) {
      if (!copy.IsPaused) {
        copy.Pause();
        RunTimer.Stop();
        PauseTimer.Start();
        PauseButton.Content = "Resume";
      }
      else {
        copy.Resume();
        PauseTimer.Stop();
        RunTimer.Start();
        PauseButton.Content = "Pause";
      }
    }

    void SourceBrowseButton_Click(object sender, RoutedEventArgs e) {
      // https://wellsr.com/vba/2016/excel/vba-select-folder-with-msoFileDialogFolderPicker/
      // https://bytes.com/topic/access/insights/916710-select-file-folder-using-filedialog-object
      if (ExcelApp == null) {
        ExcelApp = new Excel.Application();
      }
      var fileDialog = ExcelApp.get_FileDialog
        (Office.MsoFileDialogType.msoFileDialogFolderPicker);
      fileDialog.InitialFileName = Source.Text + "\\"; 
      fileDialog.InitialView = Office.MsoFileDialogView.msoFileDialogViewDetails;
      fileDialog.ButtonName = "Select Folder";
      fileDialog.Title = "Select Folder";
      fileDialog.AllowMultiSelect = true;
      int result = fileDialog.Show();
      if (result == -1 && fileDialog.SelectedItems.Count > 0) {
        Source.Text = fileDialog.SelectedItems.Cast<string>().ToArray()[0];
      }
    }

    void DestinationBrowseButton_Click(object sender, RoutedEventArgs e) {
      if (ExcelApp == null) {
        ExcelApp = new Excel.Application();
      }
      var fileDialog = ExcelApp.get_FileDialog
        (Office.MsoFileDialogType.msoFileDialogFolderPicker);
      fileDialog.InitialFileName = Destination.Text + "\\";
      fileDialog.InitialView = Office.MsoFileDialogView.msoFileDialogViewDetails;
      fileDialog.ButtonName = "Select Folder";
      fileDialog.Title = "Select Folder";
      int result = fileDialog.Show();
      if (result == -1 && fileDialog.SelectedItems.Count > 0) {
        Destination.Text = fileDialog.SelectedItems.Cast<string>().ToArray()[0];
      }
    }

    void CancelButton_Click(object sender, RoutedEventArgs e) {
      if (backupIsRunning) {
        if (prepareIsRunning == false) {
          copy.Stop();
        }
        backupIsRunning = false;
        prepareIsRunning = false;
        RunTimer.Stop();
        var ts = RunTimer.Elapsed;
        Debug.WriteLine($"Job cancelled {ts}");
        MessageBoxResult result = MessageBox.Show("Backup was cancelled",
          "Backup", MessageBoxButton.OK);
        CancelButton.IsEnabled = false;
        PauseButton.IsEnabled = false;
        BackupButton.IsEnabled = true;
      }
    }

    void Window_ContentRendered(object sender, EventArgs e) {
      //Source.Text = @"G:\Public\Files1";
      //Destination.Text = @"F:\Gitrepos2\Files1";
      Source.Text = @"G:\Public\Gitrepos\0misc";
      Destination.Text = @"F:\Gitrepos2\0misc";
      //Source.Text = @"G:\Public\Source";
      //Destination.Text = @"F:\Source";
      MessageBoxHelper.PrepToCenterMessageBoxOnForm(this);
    }

    void ScanThenCopy() {
      //Folders = new List<string>();
      //CurrentFolder.Text = "Counting Folders";

      RunTimer.Reset();
      PauseTimer.Stop();
      PauseTimer.Reset();
      RunTimer.Start();
      CancelButton.IsEnabled = true;
      PauseButton.IsEnabled = true;
      BackupButton.IsEnabled = false;
      source = Source.Text;
      exclude = optionsControl.ExcludeDirectories.Text;
      if (backupIsRunning) {
      }
      backupIsRunning = true;
      prepareIsRunning = true;
      FolderProgress.Maximum = 100;
      FolderProgress.Value = 0;

      Task task1 = ShowProgressAsync();

      Task.Run(() => {
        DirectoryInfo di = new DirectoryInfo(source);
        tf.TotalFolders = CountFolders(di, "*", exclude);
        if (tf.TotalFolders == -1) {
          backupIsRunning = false;
          return;
        }
        var ts = RunTimer.Elapsed;
        Debug.WriteLine($"End of PreScan {ts}");
        prepareIsRunning = false;
        RunTimer.Restart();
        Dispatcher.Invoke(() => {
          FolderProgress.Maximum = tf.TotalFolders;
          FolderProgress.Value = 0;
        });
        try {
          //Debug.WriteLine("b4 copy.Start");
          copy.Start();
          //Debug.WriteLine("After copy.Start");
        }
        catch (Exception e1) {
          var v1 = e1;
          throw;
        }
      });
    }

    private async Task ShowProgressAsync() {
      ByteRate.Reset();
      MbsMax = 20;
      if (!prepareIsRunning) {
      }
      while (backupIsRunning || cf.FileName != "") {
        await Task.Delay(200);
        var ts = RunTimer.Elapsed;
        if (prepareIsRunning) {
          await Dispatcher.BeginInvoke((Action)(() => {
            txtTotalTime.Text = $"{ts:hh\\:mm\\:ss\\.f}";
            if (FolderProgress.Value > 80) {
              FolderProgress.Value = 20;
            }
            else {
              FolderProgress.Value += 4; ;
            }
            ProgressLabel.Text = $"Preparing to copy {tf.TotalFolders} folders";
          }));
        }
        else {
          await ShowValues();
        }
      }
    }

    private async Task ShowValues() {
      var ts = RunTimer.Elapsed;
      await Dispatcher.BeginInvoke((Action)(() => {
        //Debug.WriteLine($"ShowValues1-\"{cf.FileName}\" {tf.BytesCopied} {cf.FileSize} {cf.FilePortion} {cf.FilePC}");
        //if (cf != null && tf.FolderCount != 0) {
        if (true) {
          if (cf.FileName != "") {
            Double pc1 = cf.FilePC;
            tf.BytesCopied -= cf.FilePortion;
            cf.FilePortion = Convert.ToInt64((double)cf.FileSize
              * pc1 / 100);
            tf.BytesCopied += cf.FilePortion;
            txtCurrentFile.Text = cf.FileName;
            txtFilePc.Text = $"{pc1}%";
            txtInFolder.Text = $"File {cf.FilesCopied} of {cf.FilesTotal}";
          }
          else {
            txtCurrentFile.Text = "";
            txtFilePc.Text = "";
            txtInFolder.Text = "";
          }
          //Debug.WriteLine($"ShowValues2-\"{cf.FileName}\" {tf.BytesCopied} {cf.FileSize} {cf.FilePortion} {cf.FilePC}");
          //txtMBytes.Text = $"{Math.Round((double)tf.BytesCopied / 1024.0 / 1024.0):#,##0}";
          //txtMBytesSkipped.Text = $"{Math.Round((double)tf.BytesSkipped / 1024.0 / 1024.0):#,##0}";
          txtMBytes.Text = $"{tf.BytesCopied:#,##0}";
          txtMBytesSkipped.Text = $"{tf.BytesSkipped:#,##0}";
          txtFiles.Text = $"{(tf.FilesCopied):#,##0}";
          txtFilesSkipped.Text = $"{(tf.FilesSkipped):#,##0}";
          txtFolders.Text = $"{(tf.FolderCount):#,##0}";
          FolderProgress.Value = tf.FolderCount;
          ProgressLabel.Text = $"Folder {tf.FolderCount} of {tf.TotalFolders}";
          CurrentFolder.Text = cf.FolderName;
          taskBarItemInfo.ProgressState = System.Windows.Shell.TaskbarItemProgressState.Normal;
          taskBarItemInfo.ProgressValue = (double)tf.FolderCount / (double)tf.TotalFolders;
          //if (cf.FileSize > 100 * 1024 * 1024) {
          //  txtCurrentFile.Text = cf.FileName;
          //  txtFilePc.Text = $"{((double)cf.FilePortion / (double)cf.FileSize):0.0%}";
          //  txtInFolder.Text = $"{cf.FilesCopied} of {cf.FilesTotal}";
          //}
          //else {
          //  txtCurrentFile.Text = "";
          //  txtFilePc.Text = "";
          //  txtInFolder.Text = "";
          //}
          try {
            string s1 = $"{ts:hh\\:mm\\:ss\\.f}";
            txtTotalTime.Text = $"{ts:hh\\:mm\\:ss\\.f}";
            if (tf.FolderCount != 0) {
              var remTime = ts.TotalMilliseconds / tf.FolderCount * tf.TotalFolders - ts.TotalMilliseconds;
              var ts1 = TimeSpan.FromMilliseconds(remTime);
              txtTimeLeft.Text = $"{ts1:hh\\:mm\\:ss\\.f}";
              DateTime ETA = DateTime.Now + ts1 + PauseTimer.Elapsed;
              txtETA.Text = $"{ETA:HH:mm:ss}";
              //TODO Track last 50 iterations
              double mbs = (double)tf.BytesCopied / 1024.0 / 1024.0 / ts.TotalSeconds;
              txtMBytesSec.Text = $"{mbs:#,##0.0}";
              if (mbs > MbsMax * 0.6) {
                MbsMax *= 1.2;
                MbsProgress.Maximum = MbsMax;
              }
              else if (mbs < MbsMax * 0.4) {
                MbsMax *= 0.8;
                MbsProgress.Maximum = MbsMax;
              }
              double mbs2 = ByteRate.Add(tf.BytesCopied, ts.TotalMilliseconds);
              txtMBytesSec2.Text = $"{mbs2:#,##0.0}";
              MbsProgress.Value = mbs2;
            }
          }
          catch (Exception e1) {
            throw e1;
          }
        }
      }));
    }

    private double GetBitRate(double bytesCopied, double totalMilliseconds) {
      throw new NotImplementedException();
    }

    long CountFolders(DirectoryInfo dir, string searchPattern, string exclude) {
      if (backupIsRunning == false) {
        return -1;
      }
      ++tf.TotalFolders;
      //TODO Count files if folders less than xx or limit folder depth
      //var files = dir.CountFileSystemObjects(DirectoryEnumerationOptions.Files);
      //TotalFiles += files;
      //Folders.Add(dir.FullName);
      DirectoryInfo[] strSubDirs = null;
      string errorMsg = null;
      try {
        strSubDirs = dir.GetDirectories();
      }
      catch (UnauthorizedAccessException e) {
        errorMsg = e.Message;
      }
      catch (Exception e) {
        errorMsg = e.Message;
        tf.TotalFolders = -1;
      }
      if (errorMsg != null) {
        if (tf.TotalFolders < 0) {
          MessageBoxResult result = MessageBox.Show(errorMsg
            + Environment.NewLine + " Unhandled Error", "Unhandled Error", MessageBoxButton.OK);
          return -1;
        }
        else if (tf.TotalFolders <= 1) {
          MessageBoxResult result = MessageBox.Show(errorMsg
            + Environment.NewLine + " Error on top level folder", "Access Error", MessageBoxButton.OK);
          return -1;
        }
        else {
          if (tf.ErrorCount == 0) {
            MessageBoxResult result = MessageBox.Show(errorMsg
              + " Do you wan't to continue and ignore all other errors?", "Access Error", MessageBoxButton.OKCancel);
            if (result != MessageBoxResult.OK) {
              return -1;
            }
          }
          ++tf.ErrorCount;
          return 0;
        }
      }

      foreach (DirectoryInfo item in strSubDirs) {
        System.IO.FileAttributes v1 = item.Attributes;
        if (item.Name == "$RECYCLE.BIN") {
        }
        if (item.Name == "System Volume Information") {
        }
        if ((item.Attributes & System.IO.FileAttributes.System)
          == System.IO.FileAttributes.System) {
          continue;
        }
        if ((item.Attributes & System.IO.FileAttributes.Directory)
          == System.IO.FileAttributes.Directory) {
          if (item.Name.Equals(exclude)) {
            continue;
          }
          try {
            long rcode = CountFolders(item, searchPattern, exclude);
            if (rcode == -1)
              return -1;
          }
          catch (Exception e) {
            Debug.Print(e.Message);
          }
        }
      }
      return tf.TotalFolders;
    }
  }
}