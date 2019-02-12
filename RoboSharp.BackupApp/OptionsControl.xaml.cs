using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace RoboSharp.BackupApp {
  /// <summary>
  /// Interaction logic for OptionsControl.xaml
  /// </summary>
  public partial class OptionsControl : UserControl {

    public OptionsControl() {
      InitializeComponent();

      CopySubdirectoriesIncludingEmpty.IsChecked = true;
      ExcludeDirectories.Text = "node_modules";
      VerboseOutput.IsChecked = true;
      VerboseOutput.IsEnabled = false;
      NoFileSizes.IsEnabled = false;
      NoProgress.IsEnabled = false;
      NoFileList.IsEnabled = false;
      NoDirectoryList.IsEnabled = false;
    }

    public void BuildRobocopyParameters(RoboCommand copy) {
      //Debugger.Instance.DebugMessageEvent += DebugMessage;
      //Debug.WriteLine("b4 new RoboCommand()");
      //copy = new RoboCommand();
      //Debug.WriteLine("after new RoboCommand()");

      // copy options
      copy.CopyOptions.FileFilter = FileFilter.Text;
      copy.CopyOptions.CopySubdirectories = CopySubDirectories.IsChecked ?? false;
      copy.CopyOptions.CopySubdirectoriesIncludingEmpty = CopySubdirectoriesIncludingEmpty.IsChecked ?? false;
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
      //Debug.WriteLine($"{DateTime.Now} before copy.Start");
      //copy.Start();
      //Debug.WriteLine($"{DateTime.Now} After copy.Start");
      //var v6 = copy.CommandOptions;
      //return copy;
      Parameters.Text = "ROBOCOPY " + copy.CommandOptions;
    }

    void IsNumeric_PreviewTextInput(object sender, TextCompositionEventArgs e) {
      e.Handled = !IsInt(e.Text);
    }

    void IsAttribute_PreviewTextInput(object sender, TextCompositionEventArgs e) {
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
  }
}
