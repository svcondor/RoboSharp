using System;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;
using Alphaleonis.Win32.Filesystem;
using System.Collections.Generic;
using System.Threading.Tasks;

// https://alphafs.alphaleonis.com/
// https://github.com/svcondor/RoboSharp
// OneNote Windows - Windows - Robocopy 2

namespace RoboSharp.BackupApp {

  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window {
    Folder cf;  // Current folder info record
    Folder tf;  // Folder to total all folders
    RoboCommand copy;
    public ObservableCollection<FileError> Errors = new ObservableCollection<FileError>();
    string SourceLower;
    string DestinationLower;
    long TotalFolders, ShowFolders, FolderCount;
    long ExtraFileCount;
    int ErrorCount;
    long ExtraByteCount;
    List<string> Folders;
    //Regex regex = new Regex(@"[ ]{2,}", RegexOptions.None);
    bool totaling = false;
    bool backupIsRunning = false;
    List<Metric> metrics;
    Stopwatch RunTimer = new Stopwatch();
    Stopwatch PauseTimer = new Stopwatch();

    public MainWindow() {
      InitializeComponent();
      this.Closing += MainWindow_Closing;
      ErrorGrid.ItemsSource = Errors;
      RoboSharp.Debugger.Instance.DebugMessageEvent += DebugMessage;
    }

    void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
      if (copy != null) {
        if (backupIsRunning) {
          MessageBox.Show("Cancel the current backup before closing", "Backup");
          e.Cancel = true;
        }
        copy.Stop();
        copy.Dispose();
      }
    }

    private void BackupButton_Click(object sender, RoutedEventArgs e) {
      totaling = false;
      MetricsGrid.ItemsSource = null;
      //OptionsGrid.IsEnabled = false;
      //ProgressTab.IsSelected = true;
      //ProgressGrid.IsEnabled = true;
      CurrentFolder.Text = "Running Preliminary Scan...";
      txtMByteCount.Text = "";
      txtErrorCount.Text = "";
      txtFileCount.Text = "";
      txtFolderCount.Text = "";
      txtFilePc.Text = "";
      txtTotalTime.Text = "";
      ErrorCount = 0;
      copy = BuildRobocopyParameters();
      Parameters.Text = "ROBOCOPY " + copy.CommandOptions;
      SourceLower = Source.Text.ToLower();
      DestinationLower = Destination.Text.ToLower();
      cf = new Folder { Name = "Source.Text" };
      tf = new Folder();
      ScanThenCopy(copy);
    }

    RoboCommand BuildRobocopyParameters() {
      copy = new RoboCommand();

      copy.OnCommandCompleted += copy_OnCommandCompleted;
      copy.OnCommandError += copy_OnCommandError;
      copy.OnCopyProgressChanged += copy_OnCopyProgressChanged;
      copy.OnError += copy_OnError;
      copy.OnFileProcessed += copy_OnFileProcessed;

      copy.CopyOptions.Source = Source.Text;
      copy.CopyOptions.Destination = Destination.Text;
      copy.CopyOptions.FileFilter = FileFilter.Text;
      copy.CopyOptions.CopySubdirectories = CopySubDirectories.IsChecked ?? false;
      copy.CopyOptions.CopySubdirectoriesIncludingEmpty
        = CopySubdirectoriesIncludingEmpty.IsChecked ?? false;
      if (!string.IsNullOrWhiteSpace(Depth.Text))
        copy.CopyOptions.Depth = Convert.ToInt32(Depth.Text);
      copy.CopyOptions.EnableRestartMode = EnableRestartMode.IsChecked ?? false;
      copy.CopyOptions.EnableBackupMode = EnableBackupMode.IsChecked ?? false;
      copy.CopyOptions.EnableRestartModeWithBackupFallback = EnableRestartModeWithBackupFallback.IsChecked ?? false;
      copy.CopyOptions.UseUnbufferedIo = UseUnbufferedIo.IsChecked ?? false;
      copy.CopyOptions.EnableEfsRawMode = EnableEfsRawMode.IsChecked ?? false;
      copy.CopyOptions.CopyFlags = CopyFlags.Text;
      copy.CopyOptions.CopyFilesWithSecurity = CopyFilesWithSecurity.IsChecked ?? false;
      copy.CopyOptions.CopyAll = CopyAll.IsChecked ?? false;
      copy.CopyOptions.RemoveFileInformation = RemoveFileInformation.IsChecked ?? false;
      copy.CopyOptions.FixFileSecurityOnAllFiles = FixFileSecurityOnAllFiles.IsChecked ?? false;
      copy.CopyOptions.FixFileTimesOnAllFiles = FixFileTimesOnAllFiles.IsChecked ?? false;
      copy.CopyOptions.Purge = Purge.IsChecked ?? false;
      copy.CopyOptions.Mirror = Mirror.IsChecked ?? false;
      copy.CopyOptions.MoveFiles = MoveFiles.IsChecked ?? false;
      copy.CopyOptions.MoveFilesAndDirectories = MoveFilesAndDirectories.IsChecked ?? false;
      copy.CopyOptions.AddAttributes = AddAttributes.Text;
      copy.CopyOptions.RemoveAttributes = RemoveAttributes.Text;
      copy.CopyOptions.CreateDirectoryAndFileTree = CreateDirectoryAndFileTree.IsChecked ?? false;
      copy.CopyOptions.FatFiles = FatFiles.IsChecked ?? false;
      copy.CopyOptions.TurnLongPathSupportOff = TurnLongPathSupportOff.IsChecked ?? false;
      if (!string.IsNullOrWhiteSpace(MonitorSourceChangesLimit.Text))
        copy.CopyOptions.MonitorSourceChangesLimit = Convert.ToInt32(MonitorSourceChangesLimit.Text);
      if (!string.IsNullOrWhiteSpace(MonitorSourceTimeLimit.Text))
        copy.CopyOptions.MonitorSourceTimeLimit = Convert.ToInt32(MonitorSourceTimeLimit.Text);

      // select options
      copy.SelectionOptions.OnlyCopyArchiveFiles = OnlyCopyArchiveFiles.IsChecked ?? false;
      copy.SelectionOptions.OnlyCopyArchiveFilesAndResetArchiveFlag = OnlyCopyArchiveFilesAndResetArchiveFlag.IsChecked ?? false;
      copy.SelectionOptions.IncludeAttributes = IncludeAttributes.Text;
      copy.SelectionOptions.ExcludeAttributes = ExcludeAttributes.Text;
      copy.SelectionOptions.ExcludeFiles = ExcludeFiles.Text;
      copy.SelectionOptions.ExcludeDirectories = ExcludeDirectories.Text;
      copy.SelectionOptions.ExcludeOlder = ExcludeOlder.IsChecked ?? false;
      copy.SelectionOptions.ExcludeJunctionPoints = ExcludeJunctionPoints.IsChecked ?? false;

      // retry options
      if (!string.IsNullOrWhiteSpace(RetryCount.Text))
        copy.RetryOptions.RetryCount = Convert.ToInt32(RetryCount.Text);
      if (!string.IsNullOrWhiteSpace(RetryWaitTime.Text))
        copy.RetryOptions.RetryWaitTime = Convert.ToInt32(RetryWaitTime.Text);

      // logging options
      copy.LoggingOptions.VerboseOutput = VerboseOutput.IsChecked ?? false;
      copy.LoggingOptions.NoFileSizes = NoFileSizes.IsChecked ?? false;
      copy.LoggingOptions.NoProgress = NoProgress.IsChecked ?? false;
      copy.LoggingOptions.ListOnly = ListOnly.IsChecked ?? false;
      copy.LoggingOptions.NoFileList = NoFileList.IsChecked ?? false;
      copy.LoggingOptions.NoDirectoryList = NoDirectoryList.IsChecked ?? false;
      copy.LoggingOptions.ReportExtraFiles = ReportExtraFiles.IsChecked ?? false;

      return copy;
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
      try {
        if (cf.ShowFile) {
          long newBytes = 0;
          if (e.CurrentFileProgress == 100) {
            newBytes = cf.FileSize - cf.FilePortion;
            cf.ShowFile = false;
          }
          else {
            newBytes = Convert.ToInt64((double)cf.FileSize
            * e.CurrentFileProgress / 100.0) - cf.FilePortion;
          }
          cf.FilePortion += newBytes;
          cf.BytesCopied += newBytes;
          tf.BytesCopied += newBytes;
          if (cf.ShowFile == false) {
            cf.FilePortion = 0;
          }
        }
        //Debug.WriteLine($"Progress\t{e.CurrentFileProgress}");
      }
      catch (Exception e1) {
        var v1 = e1;
        throw;
      }
    }

    void copy_OnError(object sender, ErrorEventArgs e) {
      Dispatcher.BeginInvoke((Action)(() => {
        try {
          Errors.Insert(0, new FileError { Error = e.Error });
          ErrorsTab.Header = string.Format("Errors ({0})", Errors.Count);
          ++ErrorCount;
          txtErrorCount.Text = $"{(ErrorCount):#,##0}";
        }
        catch (Exception) {
          throw;
        }
      }));
    }

    void copy_OnFileProcessed(object sender, FileProcessedEventArgs e) {
      try {
        var v1 = e.ProcessedFile;
        //Debug.WriteLine($"C-{v1.FileClass} T-{v1.FileClassType} N-{v1.Name} S-{v1.Size}");

        if (v1.FileClassType == FileClassType.NewDir) {
          if (v1.FileClass == "New Dir") {
            if (v1.Name.ToLower().StartsWith(DestinationLower)) {
            }
            else if (v1.Name.ToLower().StartsWith(SourceLower)) {
              ++FolderCount;
              cf = new Folder {
                Name = v1.Name,
                FilesTotal = v1.Size,
                FileName = "",
              };
            }
            else {
            }
          }
          else {
          }
        }
        else if (v1.FileClassType == FileClassType.File) {
          if (cf == null) {
          }
          if (v1.FileClass == "New File"
            || v1.FileClass == "Newer"
            || v1.FileClass == "modified") {
            cf.FileName = v1.Name;
            cf.FileSize = v1.Size;
            ++cf.FilesCopied;
            ++tf.FilesCopied;
            cf.ShowFile = true;
          }
          else if (v1.FileClass == "same"
            || v1.FileClass == "Older") {
            if (v1.FileClass == "Older") {
            }
            cf.FileSize = v1.Size;
            ++cf.FilesSkipped;
            ++tf.FilesSkipped;
            cf.BytesSkipped += v1.Size;
            tf.BytesSkipped += v1.Size;
          }
          else if (v1.FileClass == "*EXTRA File") {
            ++ExtraFileCount;
            ExtraByteCount += v1.Size;
          }
          else {
          }
        }
        else if (v1.FileClassType == FileClassType.SystemMessage) {
          Debug.WriteLine($"SYSTEM\t{v1.Name}");

          if (v1.Name.StartsWith("Total")) {
            totaling = true;
            metrics = new List<Metric>();
          }
          else if (totaling) {
            var splits = v1.Name.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (splits[1] == ":") {
              var m1 = new Metric { Type = splits[0] };
              if (splits.Length > 2) m1.Total = splits[2];
              if (splits.Length > 3) m1.Copied = splits[3];
              if (splits.Length > 4) m1.Skipped = splits[4];
              if (splits.Length > 5) m1.Mismatch = splits[5];
              if (splits.Length > 6) m1.FAILED = splits[6];
              if (splits.Length > 7) m1.Extras = splits[7];
              if (splits.Length > 8) {
              }
              metrics.Add(m1);
            }
          }
        }
        else {
        }
      }
      catch (Exception) {
        throw;
      }
    }

    void copy_OnCommandCompleted(object sender, RoboCommandCompletedEventArgs e) {
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

        txtMByteCount.Text = $"{(tf.BytesCopied / 1024 / 1024):#,##0}";
        txtErrorCount.Text = "";
        txtFileCount.Text = $"{(tf.FilesCopied + tf.FilesSkipped):#,##0}";
        txtFolderCount.Text = $"{(FolderCount):#,##0}";
        CurrentFolder.Text = cf.Name;
      }));




      Debug.WriteLine($"EVENT\tcopy_OnCommandCompleted");
      //showProgress("End Of Job", copy.counters);
      //Dispatcher.BeginInvoke((Action)(() => {
      //  //OptionsGrid.IsEnabled = true;
      //  //ProgressGrid.IsEnabled = false;
      //}));
    }

    private void PauseButton_Click(object sender, RoutedEventArgs e) {
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

    private void SourceBrowseButton_Click(object sender, RoutedEventArgs e) {
      var dialog = new System.Windows.Forms.FolderBrowserDialog();
      System.Windows.Forms.DialogResult result = dialog.ShowDialog();
      Source.Text = dialog.SelectedPath;
    }

    private void DestinationBrowseButton_Click(object sender, RoutedEventArgs e) {
      var dialog = new System.Windows.Forms.FolderBrowserDialog();
      System.Windows.Forms.DialogResult result = dialog.ShowDialog();
      Destination.Text = dialog.SelectedPath;
    }

    private void IsNumeric_PreviewTextInput(object sender, TextCompositionEventArgs e) {
      e.Handled = !IsInt(e.Text);
    }

    private void IsAttribute_PreviewTextInput(object sender, TextCompositionEventArgs e) {
      if (!Regex.IsMatch(e.Text, @"^[a-zA-Z]+$"))
        e.Handled = true;
      if ("bcefghijklmnpqrvwxyzBCEFGHIJKLMNPQRVWXYZ".Contains(e.Text))
        e.Handled = true;
      if (((TextBox)sender).Text.Contains(e.Text))
        e.Handled = true;
    }

    public static bool IsInt(string text) {
      Regex regex = new Regex("[^0-9]+$");
      return !regex.IsMatch(text);
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e) {
      copy.Stop();
      MessageBoxResult result = MessageBox.Show("Backup was cancelled",
        "Backup", MessageBoxButton.OK);
    }

    private void Window_ContentRendered(object sender, EventArgs e) {
      Source.Text = @"G:\Public\Gitrepos\0misc";
      Destination.Text = @"F:\Gitrepos2\0misc";
      CopySubdirectoriesIncludingEmpty.IsChecked = true;
      ExcludeDirectories.Text = "node_modules";
      VerboseOutput.IsChecked = true;
      MessageBoxHelper.PrepToCenterMessageBoxOnForm(this);
    }

    private void ScanThenCopy(RoboCommand copy) {
      Folders = new List<string>();
      ExtraFileCount = 0;
      ExtraByteCount = 0;
      FolderCount = 0;
      TotalFolders = 0;
      ShowFolders = 0;
      ErrorCount = 0;
      Folders = new List<string>();
      CurrentFolder.Text = "Counting Folders";
      backupIsRunning = true;

      RunTimer.Reset();
      PauseTimer.Stop();
      PauseTimer.Reset();
      RunTimer.Start();
      CancelButton.IsEnabled = true;
      PauseButton.IsEnabled = true;
      BackupButton.IsEnabled = false;

      string source = Source.Text;
      string exclude = ExcludeDirectories.Text;
      Task.Run(async () => {
        DirectoryInfo di = new DirectoryInfo(source);
        TotalFolders = CountFolders(di, "*", exclude);
        if (TotalFolders == -1) {
          StopCurrentBackup();
          return;
        }
        var ts = RunTimer.Elapsed;
        Debug.WriteLine($"End of PreScan {ts}");
        Dispatcher.Invoke(() => {
          FolderProgress.Maximum = TotalFolders;
        });
        try {
          Task t1 = copy.Start();
        }
        catch (Exception e1) {
          var v1 = e1;
          throw;
        }
        while (backupIsRunning) {
          await Task.Delay(200);
          await Dispatcher.BeginInvoke((Action)(() => {
            txtMByteCount.Text = $"{(tf.BytesCopied / 1024 / 1024):#,##0}";
            txtErrorCount.Text = "";
            txtFileCount.Text = $"{(tf.FilesCopied):#,##0}";
            txtFolderCount.Text = $"{(FolderCount):#,##0}";
            FolderProgress.Value = FolderCount;
            if (cf != null && FolderCount != 0) {
              CurrentFolder.Text = cf.Name;
              taskBarItemInfo.ProgressState = System.Windows.Shell.TaskbarItemProgressState.Normal;
              taskBarItemInfo.ProgressValue = (double)FolderCount / (double)TotalFolders;
              if (cf.FileSize > 100 * 1024) {
                txtCurrentFile.Text = cf.FileName;
                txtFilePc.Text = $"{((double)cf.FilePortion / (double)cf.FileSize):0.0%}";
                txtInFolder.Text = $"{cf.FilesCopied} of {cf.FilesTotal}";
              }
              else {
                txtCurrentFile.Text = "";
                txtFilePc.Text = "";
                txtInFolder.Text = "";
              }
              try {
                ts = RunTimer.Elapsed;
                string s1 = $"{ts:hh\\:mm\\:ss\\.f}";
                txtTotalTime.Text = $"{ts:hh\\:mm\\:ss\\.f}";
                if (FolderCount != 0) {
                  var remTime = ts.TotalMilliseconds / FolderCount * TotalFolders - ts.TotalMilliseconds;
                  var ts1 = TimeSpan.FromMilliseconds(remTime);
                  txtTimeLeft.Text = $"{ts1:hh\\:mm\\:ss\\.f}";
                  DateTime ETA = DateTime.Now + ts1 + PauseTimer.Elapsed;
                  txtETA.Text = $"{ETA:HH:mm:ss}";
                }
              }
              catch (Exception e1) {
                var v1 = e1;
                throw;
              }
            }
          }));
        }
        RunTimer.Stop();
        ts = RunTimer.Elapsed;
        Debug.WriteLine($"End of Job {ts}");
      });
    }

    private void StopCurrentBackup() {
      throw new NotImplementedException();
    }

    private long CountFolders(DirectoryInfo dir, string searchPattern, string exclude) {
      if (backupIsRunning == false) {
        return -1;
      }
      ++TotalFolders;
      if (TotalFolders >= ShowFolders + 100) {
        ShowFolders = TotalFolders;
        Dispatcher.Invoke(() => {
          CurrentFolder.Text = $"Counting Folders {ShowFolders}";
        });
      }
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
        TotalFolders = -1;
      }
      if (errorMsg != null) {
        if (TotalFolders < 0) {
          MessageBoxResult result = MessageBox.Show(errorMsg
            + Environment.NewLine + " Unhandled Error", "Unhandled Error", MessageBoxButton.OK);
          return -1;
        }
        else if (TotalFolders <= 1) {
          MessageBoxResult result = MessageBox.Show(errorMsg
            + Environment.NewLine + " Error on top level folder", "Access Error", MessageBoxButton.OK);
          return -1;
        }
        else {
          if (ErrorCount == 0) {
            MessageBoxResult result = MessageBox.Show(errorMsg
              + " Do you wan't to continue and ignore all other errors?", "Access Error", MessageBoxButton.OKCancel);
            if (result != MessageBoxResult.OK) {
              return -1;
            }
          }
          ++ErrorCount;
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
      return TotalFolders;
    }
  }
}

