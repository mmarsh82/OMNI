using OMNI.Commands;
using OMNI.Enumerations;
using OMNI.Helpers;
using OMNI.Models;
using OMNI.QMS.Enumeration;
using OMNI.QMS.Model;
using OMNI.ViewModels;
using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Input;

namespace OMNI.QMS.ViewModel
{
    class QIRFormViewModel : ViewModelBase
    {
        #region Properties

        public QIR Qir { get; set; }
        public ObservableCollection<LinkedForms> FormLinks { get; set; }
        private string selectedDisposition;
        public string SelectedDisposition
        {
            get { return selectedDisposition; }
            set
            {
                selectedDisposition = value;
                if (!string.IsNullOrEmpty(value) && value != Qir.CurrentRevision.Disposition?.Description)
                {
                    Qir.CurrentRevision.Disposition = Qir.DispositionList.Find(o => o.Description == value);
                }
                OnPropertyChanged(nameof(SelectedDisposition));
            }
        }
        private bool readOnly;
        public bool ReadOnly
        {
            get
            { return readOnly; }
            set
            {
                readOnly = value;
                OnPropertyChanged(nameof(ReadOnly));
            }
        }
        private int? tempID;
        public int? TempID
        {
            get
            {
                return Qir?.IDNumber == null ? tempID : Qir.IDNumber;
            }
            set
            {
                if (value != Qir.IDNumber && value != null)
                {
                    Qir = new QIR(value, false);
                    OnPropertyChanged(nameof(Qir));
                    tempID = Qir?.IDNumber == null ? value : Qir.IDNumber;
                    OnPropertyChanged(nameof(TempID));
                    SelectedDisposition = Qir.CurrentRevision?.Disposition?.Description;
                    OnPropertyChanged(nameof(Lot));
                    if (Qir.LinkExists())
                    {
                        var _remove = Qir.GetLinkList();
                    }
                }
                else
                {
                    tempID = value;
                    OnPropertyChanged(nameof(TempID));
                }
            }
        }
        private FormCommand commandType;
        public FormCommand CommandType
        {
            get
            {
                return commandType;
            }
            set
            {
                commandType = value;
                OnPropertyChanged(nameof(CommandType));
            }
        }
        public string Lot
        {
            get { return Qir.CurrentRevision?.LotNumber; }
            set
            {
                value = value.ToUpper();
                if (!string.IsNullOrEmpty(value) && Qir.LoadM2kData && value != "N/A")
                {
                    Qir.LoadM2kData = false;
                    Qir.GetQIRFromM2k(value, M2kDataQuery.LotNumber);
                    Qir.LoadM2kData = true;
                }
                Qir.CurrentRevision.LotNumber = value;
                OnPropertyChanged(nameof(Lot));
            }
        }

        public ObservableCollection<string> PicCollection { get; set; }

        RelayCommand _formCommand;
        RelayCommand _attachCommand;
        RelayCommand _printCommand;
        RelayCommand _emailCommmand;
        RelayCommand _viewPhotoCommand;

        #endregion

        /// <summary>
        /// QIR Form ViewModel Constructor
        /// </summary>
        public QIRFormViewModel()
        {
            CommandType = FormCommand.Submit;
            Qir = new QIR();
            TempID = Qir.IDNumber;
            ReadOnly = CurrentUser.Quality;
            if (FormLinks == null && Qir.FormLinkList != null)
            {
                FormLinks = new ObservableCollection<LinkedForms>(Qir.FormLinkList);
            }
            PicCollection = new ObservableCollection<string>(Users.GetQirPICIList());
        }

        /// <summary>
        /// QIR Form ViewModel Constructor
        /// <param name="qirEZ">Set to QIR EZ Type</param>
        /// </summary>
        public QIRFormViewModel(bool qirEZ)
        {
            CommandType = FormCommand.Submit;
            Qir = new QIR(qirEZ);
            TempID = Qir.IDNumber;
            ReadOnly = true;
            if (FormLinks == null && Qir.FormLinkList != null)
            {
                FormLinks = new ObservableCollection<LinkedForms>(Qir.FormLinkList);
            }
            Qir.PIC = CurrentUser.GetFullName();
        }

        /// <summary>
        /// QIR Form ViewModel Constructor
        /// </summary>
        /// <param name="qir">QIR Object to load</param>
        public QIRFormViewModel(QIR qir)
        {
            CommandType = FormCommand.Update;
            Qir = qir;
            SelectedDisposition = Qir.CurrentRevision.Disposition.Description;
            TempID = qir.IDNumber;
            ReadOnly = CurrentUser.Quality;
            if (Qir.LinkExists())
            {
                var _remove = Qir.GetLinkList();
            }
            if (FormLinks == null && Qir.FormLinkList != null)
            {
                FormLinks = new ObservableCollection<LinkedForms>(Qir.FormLinkList);
            }
            PicCollection = new ObservableCollection<string>(Users.GetQirPICIList());
            if (!PicCollection.Contains(Qir.PIC) && !string.IsNullOrEmpty(Qir.PIC))
            {
                PicCollection.Add(Qir.PIC);
            }
        }

        /// <summary>
        /// TODO: move to its own usercontrol
        /// </summary>
        public void UpdateUILinkList()
        {
            if (Qir != null)
            {
                FormLinks = new ObservableCollection<LinkedForms>(Qir.FormLinkList);
                OnPropertyChanged(nameof(FormLinks));
            }
        }

        #region FormICommand Implementation

        public ICommand FormICommand
        {
            get
            {
                if (_formCommand == null)
                {
                    _formCommand = new RelayCommand(FormCommandExecute, FormCommandCanExecute);
                }
                return _formCommand;
            }
        }
        public void FormCommandExecute(object parameter)
        {
            parameter.ToString();
            switch(CommandType)
            {
                case FormCommand.Submit:
                    Qir.IDNumber = Qir.Submit(Qir);
                    TempID = Qir.IDNumber;
                    OnPropertyChanged(nameof(Qir));
                    SelectedDisposition = Qir.CurrentRevision.Disposition.Description;
                    CommandType = FormCommand.Update;
                    break;
                case FormCommand.Update:
                    Qir.Update(Qir);
                    break;
            }
        }
        public bool FormCommandCanExecute(object parameter) => 
            ReadOnly && parameter != null
            && Qir != null
            && (Qir.CurrentRevision?.Shift >= 1
            && Qir.CurrentRevision?.Shift <= 3)
            && !string.IsNullOrEmpty(Qir.WONumber)
            && !string.IsNullOrEmpty(Qir.CurrentRevision?.LotNumber)
            && Qir.CurrentRevision?.NCMCode > 0
            && Qir.CurrentRevision?.Origin > 0
            && Qir.CurrentRevision?.MaterialLost >= 0
            && Qir.Found > 0
            && !string.IsNullOrEmpty(Qir.CurrentRevision?.Cause)
            && Qir.CurrentRevision?.Disposition != null
            && !string.IsNullOrEmpty(Qir.PIC);

        #endregion

        #region AttachPhotoICommand Implementation

        public ICommand AttachPhotoICommand
        {
            get
            {
                if (_attachCommand == null)
                {
                    _attachCommand = new RelayCommand(AttachPhotoExecute, AttachPhotoCanExecute);
                }
                return _attachCommand;
            }
        }
        private void AttachPhotoExecute(object parameter)
        {
            if (!Qir.AttachPhoto())
            {
                ExceptionWindow.Show("Attachment failure", "OMNI is not able to attach photos to this QIR at this time.\nPlease contact IT right away to trouble shoot this issue.");
            }
            else
            {
                OnPropertyChanged(nameof(Qir));
            }
        }
        private bool AttachPhotoCanExecute(object parameter) => ReadOnly && Qir?.IDNumber != null;

        #endregion

        #region PrintFormICommand Implementation

        public ICommand PrintFormICommand
        {
            get
            {
                if (_printCommand == null)
                {
                    _printCommand = new RelayCommand(PrintFormExecute, PrintFormCanExecute);
                }
                return _printCommand;
            }
        }
        private void PrintFormExecute(object parameter)
        {
            new QIR(Convert.ToInt32(Qir.IDNumber), true).ExportToPDF(true);
            PrintForm.FromPDF($"{Properties.Settings.Default.omnitemp}{Qir.IDNumber.ToString()}.pdf");
            if (File.Exists($"{Properties.Settings.Default.omnitemp}{Qir.IDNumber.ToString()}.pdf"))
            {
                File.Delete($"{Properties.Settings.Default.omnitemp}{Qir.IDNumber.ToString()}.pdf");
            }
        }
        private bool PrintFormCanExecute(object parameter) => Qir?.IDNumber != null && App.SqlConAsync.State == System.Data.ConnectionState.Open;

        #endregion

        #region EmailFormICommand Implementation

        public ICommand EmailFormICommand
        {
            get
            {
                if (_emailCommmand == null)
                {
                    _emailCommmand = new RelayCommand(EmailFormExecute, EmailFormCanExecute);
                }
                return _emailCommmand;
            }
        }
        private void EmailFormExecute(object parameter)
        {
            try
            {
                Qir.ExportToPDF(true);
                if (File.Exists($"{Properties.Settings.Default.QIRPhotoDirectory}{Qir.IDNumber}P.docx"))
                {
                    EmailForm.ManualSend($"QIR {Qir.IDNumber}", $"{Properties.Settings.Default.omnitemp}{Qir.IDNumber}.pdf", $"{Properties.Settings.Default.QIRPhotoDirectory}{Qir.IDNumber}P.docx");
                }
                else
                {
                    EmailForm.ManualSend($"QIR {Qir.IDNumber}", $"{Properties.Settings.Default.omnitemp}{Qir.IDNumber}.pdf");
                }
                if (File.Exists($"{Properties.Settings.Default.omnitemp}{Qir.IDNumber}.pdf"))
                {
                    File.Delete($"{Properties.Settings.Default.omnitemp}{Qir.IDNumber}.pdf");
                }
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
            }
        }
        private bool EmailFormCanExecute(object parameter) => Qir?.IDNumber != null;

        #endregion

        #region ViewPhotoICommand Implementation

        public ICommand ViewPhotoICommand
        {
            get
            {
                if (_viewPhotoCommand == null)
                {
                    _viewPhotoCommand = new RelayCommand(ViewPhotoExecute);
                }
                return _viewPhotoCommand;
            }
        }
        private void ViewPhotoExecute(object parameter)
        {
            Process.Start($"{Properties.Settings.Default.QIRPhotoDirectory}{Qir.IDNumber}P.docx");
        }

        #endregion

        /// <summary>
        /// Object disposal
        /// </summary>
        /// <param name="disposing">Called by the GC Finalizer</param>
        public override void OnDispose(bool disposing)
        {
            if (disposing)
            {
                Qir = null;
                _attachCommand = _emailCommmand = _formCommand =_printCommand = null;
            }
        }
    }
}
