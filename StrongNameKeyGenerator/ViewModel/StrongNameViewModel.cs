using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using StrongNameKeyGenerator.View;
using System.Text.RegularExpressions;

namespace StrongNameKeyGenerator.ViewModel
{
    public class StrongNameViewModel : INotifyPropertyChanged
    {
        private bool _isProcessing = false;
        private bool _isError = false;
        private StrongNameWindow dialog;

                public StrongNameViewModel()
        {
           
        }


        #region Public properties

      
            private string _message;
        public string Message
        {
            get
            {
                return _message;
            }

            set
            {
                if (value != _message)
                {
                    _message = value;
                    OnPropertyChanged();
                }
            }
        }


        private string _outFileName="Benjamin";
        public string OutFileName
        {
            get
            {
                return  _outFileName ;
            }

            set
            {
                if (value != _outFileName)
                {
                    _outFileName = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _outFileKey;
        public string OutFileKey
        {
            get
            {
                return  _outFileKey;
            }

            set
            {
                if (value != _outFileKey)
                {
                    _outFileKey = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _strongKey;
        public string StrongKey
        {
            get
            {
                return _strongKey;
            }

            set
            {
                if (value != _strongKey)
                {
                    _strongKey = value;
                    OnPropertyChanged();
                }
            }
        }

        private ICommand _okCommand;
        public ICommand OkCommand
        {
            get { return _okCommand ?? (_okCommand = new RelayCommand(CloseDialogWindow)); }
        }

        public EnvDTE80.DTE2 DTE { get; internal set; }
        public string FilePath { get; internal set; }
        public string Argument { get; internal set; }
        public string ProjectDirectory { get; internal set; }
        public string StrongName { get; internal set; }


        #endregion
        
        private void CloseDialogWindow()
        {
            dialog.Close();
            dialog = null;
        }

        internal async void StrongNamekeyGeneration()
        {
            if (_isProcessing)
                return;

            await Dispatcher.CurrentDispatcher.BeginInvoke(new Action(async () =>
            {
                await ExecuteShellCommand();

            }), DispatcherPriority.SystemIdle, null);

        }

        private async Task ExecuteShellCommand()
        {
            var hwnd = new IntPtr(DTE.MainWindow.HWnd);
            var window = (Window)System.Windows.Interop.HwndSource.FromHwnd(hwnd).RootVisual;

            dialog = new StrongNameWindow
            {
                Owner = window,
                DataContext = this
            };
            dialog.Show();


            await Task.Run(() =>
            {
                StartStrongNamekeyGeneration ();
            });

        }

             private void StartStrongNamekeyGeneration()
        {

            var commandDirective = "/k ";
            var command = "\"" + FilePath + "\"";

            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe",
                Arguments = commandDirective +  "\"" + command + "\"",
                WorkingDirectory = ProjectDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true
            };

            try
            {

                using (System.Diagnostics.Process p = new System.Diagnostics.Process())
                {
                    var exeProcess = System.Diagnostics.Process.Start(startInfo);
                    exeProcess.EnableRaisingEvents = true;

                    exeProcess.OutputDataReceived += OutputDataReceived;

                    if (exeProcess != null)
                    {

                        //var generateStrongKey = @"sn -p key.snk " + OutFileName + ".key";
                        //exeProcess.StandardInput.WriteLine(generateStrongKey);
                        //exeProcess.StandardInput.WriteLine(@"sn -t " + OutFileName + ".key");


                        exeProcess.StandardInput.WriteLine(@"sn -p " + StrongName + " " + OutFileName + ".key");
                        exeProcess.StandardInput.WriteLine(@"sn -t " + OutFileName + ".key");


                        //exeProcess.StandardInput.WriteLine(@"sn -p key.snk outfileNever5.key");
                        //exeProcess.StandardInput.WriteLine(@"sn -t outfileNever5.key");



                        exeProcess.StandardInput.Close();

                        exeProcess.BeginOutputReadLine();

                        exeProcess.WaitForExit();
                    }
                }
            }
            catch (Exception ex)
            {
                _isError = true;
                _isProcessing = false;
                Logger.Log(ex.Message.ToString());
                Message = ex.Message.ToString();
            }
        }

          private void OutputDataReceived(object sender, System.Diagnostics.DataReceivedEventArgs e)
        {
            if (e == null || string.IsNullOrEmpty(e.Data))
                return;

            if(e.Data.Contains("Public key token is"))
            {
                StrongKey=e.Data.Replace("Public key token is ", "");
            }

            Logger.Log(e.Data);
        }

        private void ErrorDataReceived(object sender, System.Diagnostics.DataReceivedEventArgs e)
        {
            if (e == null || string.IsNullOrEmpty(e.Data))
                return;

            Logger.Log(e.Data);
            Message = e.Data;
            _isError = true;
            _isProcessing = false;

        }



        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
