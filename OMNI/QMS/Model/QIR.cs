using iTextSharp.text.pdf;
using Microsoft.Win32;
using OMNI.Extensions;
using OMNI.Helpers;
using OMNI.Models;
using OMNI.QMS.Enumeration;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace OMNI.QMS.Model
{
    public class QIR : FormBase, INotifyPropertyChanged
    {
        #region Properties

        public QIRType QIRFormType { get; set; }
        private string partNumber;
        public string PartNumber
        {
            get { return partNumber; }
            set
            {
                value = value?.ToUpper();
                if ((CurrentRevision.LotNumber == "N/A" || string.IsNullOrEmpty(CurrentRevision.LotNumber)) && (WONumber == "N/A" || string.IsNullOrEmpty(WONumber)) && !string.IsNullOrEmpty(value) && LoadM2kData)
                {
                    LoadM2kData = false;
                    CurrentRevision.LotNumber = WONumber = "N/A";
                    OnPropertyChanged(nameof(CurrentRevision));
                    this.GetQIRFromM2k(value, M2kDataQuery.PartNumber);
                    LoadM2kData = true;
                }
                partNumber = value;
                OnPropertyChanged(nameof(PartNumber));
            }
        }
        private string woNumber;
        public string WONumber
        {
            get { return woNumber; }
            set
            {
                value = value.ToUpper();
                if (!string.IsNullOrEmpty(value) && value != woNumber && value.Length == 6 && LoadM2kData)
                {
                    LoadM2kData = false;
                    this.GetQIRFromM2k(value, M2kDataQuery.WONumber);
                    LoadM2kData = true;
                }
                woNumber = value;
                OnPropertyChanged(nameof(WONumber));
            }
        }
        public List<WorkCenter> WorkCenterList { get { return WorkCenter.GetListAsync(Enumerations.WorkCenterType.QMS).Result; } }
        private int found;
        public int Found
        {
            get
            { return found; }
            set
            { found = value; OnPropertyChanged(nameof(Found)); }
        }
        private List<NCM> ncmCodeList;
        public List<NCM> NCMCodeList
        {
            get { return ncmCodeList; }
            set
            {
                ncmCodeList = QIRFormType == QIRType.QIR || !IsNew ? NCM.GetNCMListAsync().Result : NCM.GetNCMListAsync(Found).Result;
                OnPropertyChanged(nameof(NCMCodeList));
            }
        }
        private double materialCost;
        public double MaterialCost { get { return materialCost; } set { materialCost = value; OnPropertyChanged(nameof(MaterialCost)); } }
        private string uom;
        public string UOM { get { return uom; } set { uom = value; OnPropertyChanged(nameof(UOM)); } }
        private List<string> causeList;
        public List<string> CauseList
        { 
            get { return causeList; }
            set { causeList = GetQIRCauseListAsync(); OnPropertyChanged(nameof(CauseList)); }
        }
        public List<Supplier> SupplierList { get { return Supplier.GetSupplierListAsync().Result; } }
        public List<QIRDisposition> DispositionList { get { return QIRDisposition.GetQIRDispositionListAsync().Result; } }
        public string PIC { get; set; }
        public ObservableCollection<QIRRevision> RevisionList { get; set; }
        private QIRRevision currentRevision;
        public QIRRevision CurrentRevision
        {
            get { return currentRevision; }
            set { currentRevision = value; OnPropertyChanged(nameof(CurrentRevision)); if (value != null) { Lost = null; } }
        }
        public bool Escape { get { return Found == CurrentRevision?.Origin; } }
        private bool isPhotosAttached;
        public bool IsPhotosAttached
        {
            get
            {
                return isPhotosAttached;
            }
            set
            {
                if (IDNumber != null && value)
                {
                    value = Directory.GetFiles(Properties.Settings.Default.QIRPhotoDirectory, $"{IDNumber}P.*", SearchOption.TopDirectoryOnly).Length > 0;
                }
                if (value != isPhotosAttached && !value)
                {
                    var _result = System.Windows.MessageBox.Show("Are you sure that you want to delete there pictures?", "Deletion Validation",
                        System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning, System.Windows.MessageBoxResult.No);
                    if (_result == System.Windows.MessageBoxResult.Yes)
                    {
                        File.Delete($"{Properties.Settings.Default.QIRPhotoDirectory}{IDNumber}P.docx");
                    }
                    else
                    {
                        value = true;
                    }
                }
                isPhotosAttached = value;
                OnPropertyChanged(nameof(IsPhotosAttached));
            }
        }
        private int? lost;
        public int? Lost
        {
            get { return lost; }
            set
            {
                lost = value ?? CurrentRevision.MaterialLost;
                CurrentRevision.MaterialLost = Convert.ToInt32(lost);
                OnPropertyChanged(nameof(Lost));
                OnPropertyChanged(nameof(Total));
            }
        }
        public double Total { get { return Convert.ToInt32(Lost) * MaterialCost; } }
        public bool LoadM2kData;
        private readonly bool IsNew = true;
        public static DateTime LastUpdate;

        public bool legacy;
        public bool IsLegacy
        {
            get
            { return legacy; }
            set
            {
                legacy = value;
                OnPropertyChanged(nameof(IsLegacy));
            }
        }

        #endregion

        #region INotifyPropertyChanged Implementation

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null)
            {
                var e = new PropertyChangedEventArgs(propertyName);
                handler(this, e);
            }
        }

        #endregion

        /// <summary>
        /// QIR Object Constructor
        /// </summary>
        public QIR()
        {
            FormModule = Enumerations.Module.QIR;
            QIRFormType = QIRType.QIR;
            IDNumber = 0;
            Date = DateTime.Now;
            Submitter = CurrentUser.FullName;
            RevisionList = new ObservableCollection<QIRRevision>(new List<QIRRevision>());
            CurrentRevision = new QIRRevision(false);
            LoadM2kData = true;
            NCMCodeList = null;
            CauseList = null;
            IsLegacy = false;
        }

        /// <summary>
        /// QIR Object Constructor
        /// <param name="qirEZ">Set to QIR EZ type</param>
        /// </summary>
        public QIR(bool qirEZ)
        {
            FormModule = Enumerations.Module.QIR;
            QIRFormType = QIRType.QIREZ;
            IDNumber = 0;
            Date = DateTime.Now;
            Submitter = PIC = CurrentUser.FullName;
            CurrentRevision = new QIRRevision(qirEZ);
            RevisionList = new ObservableCollection<QIRRevision>(new List<QIRRevision>());
            LoadM2kData = true;
        }

        /// <summary>
        /// QIR Object Constructor
        /// </summary>
        /// <param name="qirNumber">QIR Object ID Number to load</param>
        /// <param name="validate">Validate that the QIR ID Number exists in the OMNI DataBase</param>
        public QIR(int? qirNumber, bool validate)
        {
            LoadM2kData = false;
            IsNew = false;
            RevisionList = new ObservableCollection<QIRRevision>(QIRRevision.GetQIRRevisionList(qirNumber, DispositionList));
            if (RevisionList.Count > 0)
            {
                CurrentRevision = RevisionList[0];
            }
            FormModule = Enumerations.Module.QIR;
            try
            {
                if (validate)
                {
                    using (SqlCommand cmd = new SqlCommand($"USE {App.DataBase}; SELECT COUNT([QIRNumber]) FROM [qir_master] WHERE [QIRNumber]=@p1", App.SqlConAsync))
                    {
                        cmd.Parameters.AddWithValue("@p1", qirNumber);
                        if (Convert.ToInt32(cmd.ExecuteScalar()) == 0)
                        {
                            ExceptionWindow.Show("Invalid QIR Number", $"{qirNumber} is invalid.\nPlease double check your entry and try again.");
                            return;
                        }
                    }
                }
                using (SqlCommand cmd = new SqlCommand($"SELECT * FROM [qir_master] WHERE [QIRNumber]=@p1", App.SqlConAsync))
                {
                    cmd.Parameters.AddWithValue("@p1", qirNumber);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            IDNumber = qirNumber;
                            Date = reader.SafeGetDateTime("QIRDate");
                            Submitter = reader.SafeGetString(nameof(Submitter));
                            QIRFormType = (QIRType)Enum.Parse(typeof(QIRType), reader.SafeGetString("Type").Replace(" ", string.Empty));
                            PartNumber = !reader.IsDBNull(4) ? reader.SafeGetString(nameof(PartNumber)) : string.Empty;
                            WONumber = !reader.IsDBNull(5) ? reader.SafeGetString(nameof(WONumber)) : string.Empty;
                            Found = !reader.IsDBNull(6) ? reader.SafeGetInt32(nameof(Found)) : 0;
                            MaterialCost = !reader.IsDBNull(7) ? reader.SafeGetDouble(nameof(MaterialCost)) : 0.00;
                            UOM = !reader.IsDBNull(9) ? reader.SafeGetString(nameof(UOM)) : string.Empty;
                            PIC = !reader.IsDBNull(10) ? reader.SafeGetString(nameof(PIC)) : string.Empty;
                        }
                    }
                }
                NotesTable = this.GetNotesTable();
                IsPhotosAttached = true;
                NCMCodeList = null;
                if (CurrentRevision != null && NCM.IsLegacyCode(CurrentRevision.NCMCode))
                {
                    NCMCodeList.Add(new NCM { Code = CurrentRevision.NCMCode, Summary = NCM.GetLegacySummary(CurrentRevision.NCMCode) });
                }
                CauseList = null;
                if (CurrentRevision != null && IsLegacyCause(CurrentRevision.Cause))
                {
                    CauseList.Add(CurrentRevision.Cause);
                }
                IsLegacy = !string.IsNullOrEmpty(CurrentRevision?.CauseReason) || !string.IsNullOrEmpty(CurrentRevision?.DispositionReason);
                LoadM2kData = true;
            }
            catch (Exception)
            {
                return;
            }
        }

        public static bool IsLegacyCause(string description)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand($"USE {App.DataBase}; SELECT COUNT([IDNumber]) FROM [qir_cause] WHERE [Description] = @p1", App.SqlConAsync))
                {
                    cmd.Parameters.AddWithValue("p1", description);
                    var _exists = int.TryParse(cmd.ExecuteScalar().ToString(), out int i) && i > 0;
                    return _exists && (description != "External Vendor Quality" || description != "Internal Vendor Quality");
                }
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
            }
            return false;
        }

        /// <summary>
        /// List of QIR Causes
        /// </summary>
        /// <returns>Generated List of QIR Causes</returns>
        public static List<string> GetQIRCauseListAsync()
        {
            return new List<string> { "External Vendor Quality", "Internal Vendor Quality" };
        }

        /// <summary>
        /// Flagged Property DataTable Column Changing Event
        /// </summary>
        /// <param name="sender">Column changing</param>
        /// <param name="e">Parent DataTable</param>
        public static void FlaggedColumnChanging(object sender, DataColumnChangeEventArgs e)
        {
            try
            {
                if (e.Column.ColumnName == "Flagged")
                {
                    using (SqlCommand cmd = new SqlCommand($"USE {App.DataBase}; UPDATE [qir_notice] SET [{CurrentUser.IdNumber}]=@p1 WHERE [QIRNumber]=@p2", App.SqlConAsync))
                    {
                        cmd.Parameters.AddWithValue("p1", e.ProposedValue);
                        cmd.Parameters.AddWithValue("p2", Convert.ToInt32(e.Row.ItemArray[2]));
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex, nameof(FlaggedColumnChanging));
            }
        }

        /// <summary>
        /// Load the QIR Notice into a DataTable
        /// </summary>
        /// <param name="site">Optional: Site to filter the Notice data</param>
        /// <param name="update">Optional: Update call as a boolean</param>
        /// <returns>Filtered Notice DataTable</returns>
        public static DataTable LoadNotice(string site = "", bool update = false)
        {
            site = site == string.Empty ? CurrentUser.Site : site;
            var table = new DataTable();
            try
            {
                var cmdString = $"USE {App.DataBase}; SELECT n.[{CurrentUser.IdNumber}] as 'Flagged', r.[SupplierID], q.* FROM [qir_master] q ";
                cmdString += $"LEFT JOIN [qir_notice] n ON q.[QIRNumber]=n.[QIRNumber] ";
                cmdString += $"LEFT JOIN [qir_revisions] r ON q.[QIRNumber]=r.[QIRNumber] AND q.[QIRDate]=r.[revision_date]";
                if (update)
                {
                    cmdString += $" WHERE [QIRDate]>'{LastUpdate.ToString("yyyy-MM-dd HH:mm")}'";
                }
                else if (site == "WCCO")
                {
                    cmdString += $" WHERE (q.[Status]= 'Open' OR n.[{ CurrentUser.IdNumber}]= 1 OR n.[{CurrentUser.IdNumber}] IS NULL) OR q.[QIRDate] BETWEEN '{DateTime.Now.AddDays(-CurrentUser.NoticeHistory).ToString("yyyy-MM-dd HH:mm:ss")}' AND '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}'";
                }
                else
                {
                    cmdString += $" WHERE ((q.[Status]= 'Open' OR n.[{ CurrentUser.IdNumber}]= 1 OR n.[{CurrentUser.IdNumber}] IS NULL) AND r.[SupplierID]=1015) OR q.[QIRDate] BETWEEN '{DateTime.Now.AddDays(-CurrentUser.NoticeHistory).ToString("yyyy-MM-dd HH:mm:ss")}' AND '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' AND r.[SupplierID]=1015";
                }
                using (var adapter = new SqlDataAdapter(cmdString, App.SqlConAsync))
                {
                    adapter.Fill(table);
                    LastUpdate = DateTime.Now;
                    return table;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Submit a QIR or QIR EZ to the QIR Master DataBase
        /// </summary>
        /// <param name="qirObject">QIR Object</param>
        /// <returns>Last inserted QIR.IDNumber</returns>
        public override int? Submit(object qirObject)
        {
            var _qir = (QIR)qirObject;
            var _idNumber = 0;
            try
            {
                using (SqlCommand cmd = new SqlCommand($@"USE {App.DataBase};
                                                        INSERT INTO
                                                            [qir_master] ([QIRDate], [Submitter], [Type], [PartNumber], [WONumber], [Found], [MaterialCost], [TotalCost], [UOM], [PIC], [Status])
                                                        Values(@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8, @p9, @p10, @p11);
                                                        SELECT [QIRNumber] FROM [qir_master] WHERE [QIRNumber] = @@IDENTITY;", App.SqlConAsync))
                {
                    cmd.Parameters.AddWithValue("p1", _qir.CurrentRevision.RevDate.ToString("yyyy-MM-dd HH:mm:ss"));
                    cmd.Parameters.AddWithValue("p2", _qir.CurrentRevision.RevSubmitter);
                    cmd.Parameters.AddWithValue("p3", _qir.QIRFormType.ToString());
                    cmd.SafeAddParameters("p4", _qir.PartNumber);
                    cmd.SafeAddParameters("p5", _qir.WONumber);
                    cmd.Parameters.AddWithValue("p6", _qir.Found);
                    cmd.Parameters.AddWithValue("p7", _qir.MaterialCost);
                    cmd.Parameters.AddWithValue("p8", _qir.MaterialCost * _qir.CurrentRevision.MaterialLost);
                    cmd.SafeAddParameters("p9", _qir.UOM);
                    cmd.SafeAddParameters("p10", _qir.PIC);
                    cmd.Parameters.AddWithValue("p11", _qir.CurrentRevision.Disposition.Status.ToString());
                    _idNumber = Convert.ToInt32(cmd.ExecuteScalar());
                }
                _qir.CurrentRevision.Submit(Convert.ToInt32(_idNumber));
                _qir.RevisionList.Add(CurrentRevision);
                return _idNumber;
            }
            catch (Exception e)
            {
                ExceptionWindow.Show("Unhandled Exception", e.Message, e);
                return null;
            }
        }

        /// <summary>
        /// Update the QIR in the QIR Master DataBase
        /// </summary>
        /// <param name="qirObject">QIR Object</param>
        public override void Update(object qirObject)
        {
            var _qir = (QIR)qirObject;
            try
            {
                var _tempRev = _qir.CurrentRevision;
                _tempRev.RevDate = DateTime.Now;
                using (SqlCommand cmd = new SqlCommand($@"USE {App.DataBase};
                                                        UPDATE
                                                            [qir_master]
                                                        SET
                                                            [QIRDate]=@p1,
                                                            [Submitter]=@p2,
                                                            [PartNumber]=@p3,
                                                            [WONumber]=@p4,
                                                            [Found]=@p5,
                                                            [MaterialCost]=@p6,
                                                            [TotalCost]=@p7,
                                                            [UOM]=@p8,
                                                            [PIC]=@p9,
                                                            [Status]=@p10
                                                        WHERE
                                                            [QIRNumber]=@p11", App.SqlConAsync))
                {
                    cmd.Parameters.AddWithValue("p1", _tempRev.RevDate.ToString("yyyy-MM-dd HH:mm:ss"));
                    cmd.Parameters.AddWithValue("p2", CurrentUser.FullName);
                    cmd.SafeAddParameters("p3", _qir.PartNumber);
                    cmd.Parameters.AddWithValue("p4", _qir.WONumber);
                    cmd.Parameters.AddWithValue("p5", _qir.Found);
                    cmd.Parameters.AddWithValue("p6", _qir.MaterialCost);
                    cmd.Parameters.AddWithValue("p7", _qir.MaterialCost * _qir.CurrentRevision.MaterialLost);
                    cmd.Parameters.AddWithValue("p8", _qir.UOM);
                    cmd.SafeAddParameters("p9", _qir.PIC);
                    cmd.Parameters.AddWithValue("p10", _qir.CurrentRevision.Disposition.Status.ToString());
                    cmd.Parameters.AddWithValue("p11", _qir.IDNumber);
                    cmd.ExecuteNonQuery();
                }
                _qir.RevisionList[_qir.RevisionList.IndexOf(_qir.CurrentRevision)] = QIRRevision.GetQIRRevision(_qir.IDNumber, CurrentRevision.RevNumber, _qir.DispositionList);
                _tempRev.Submit(Convert.ToInt32(_qir.IDNumber));
                _qir.RevisionList.Insert(0, _tempRev);
                _qir.CurrentRevision = _tempRev;

            }
            catch (Exception e)
            {
                ExceptionWindow.Show("Unhandled Exception", e.Message, e);
            }
        }

        /// <summary>
        /// Mark all items in the Notice table to viewed for current user
        /// </summary>
        public static void MarkAllViewed()
        {
            using (SqlCommand cmd = new SqlCommand($"USE {App.DataBase}; UPDATE [qir_notice] SET [{CurrentUser.IdNumber}]=0 WHERE [{CurrentUser.IdNumber}] IS NULL", App.SqlConAsync))
            {
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Get Metrics Table from OMNI Database
        /// </summary>
        /// <param name="startDate">Data retreival start date</param>
        /// <param name="endDate">Data retreival end date</param>
        /// <returns>Metrics Table as DataTable</returns>
        public static DataTable GetTableData(DateTime startDate, DateTime endDate)
        {
            using (DataTable dt = new DataTable())
            {
                try
                {
                    using (SqlDataAdapter da = new SqlDataAdapter($"USE {App.DataBase}; SELECT * FROM [qir_metrics_view] WHERE [QIRDate] BETWEEN '{startDate.ToString("yyyy-MM-dd")}' AND '{endDate.AddDays(1).ToString("yyyy-MM-dd")}'", App.SqlConAsync))
                    {
                        da.Fill(dt);
                    }
                }
                catch (Exception e)
                {
                    ExceptionWindow.Show("Unhandled Exception", e.Message, e);
                }
                return dt;
            }
        }
    }

    public static class QIRExtension
    {
        /// <summary>
        /// Export to FORM5008 rev N pdf.  This will open the created form up right after if not autosaved.
        /// </summary>
        /// <param name="qirObject">Currnet QIR Object</param>
        /// <param name="autoSave">optional: Autosave in the OMNI temp file</param>
        public static void ExportToPDF(this QIR qirObject, bool autoSave = false)
        {
            if (qirObject != null)
            {
                try
                {
                    using (PdfReader reader = new PdfReader(Properties.Settings.Default.QIRDocument))
                    {
                        using (PdfStamper stamp = new PdfStamper(reader, new FileStream($"{Properties.Settings.Default.omnitemp}{qirObject.IDNumber}.pdf", FileMode.Create)))
                        {
                            var pdfField = stamp.AcroFields;
                            switch (qirObject.CurrentRevision.Cause)
                            {
                                case "Material":
                                    pdfField.SetField("CheckBox1[6]", "1");
                                    break;
                                case "Method":
                                    pdfField.SetField("CheckBox1[7]", "1");
                                    break;
                                case "Equipment":
                                    pdfField.SetField("CheckBox1[2]", "1");
                                    break;
                                case "Miscellaneous":
                                    pdfField.SetField("CheckBox1[8]", "1");
                                    break;
                            }
                            switch (qirObject.CurrentRevision.Disposition.Description)
                            {
                                case "Reject":
                                    pdfField.SetField("CheckBox1[0]", "1");
                                    break;
                                case "Discard":
                                    pdfField.SetField("CheckBox1[1]", "1");
                                    break;
                                case "Accept":
                                    pdfField.SetField("CheckBox1[11]", "1");
                                    break;
                                case "Quality Hold":
                                    pdfField.SetField("CheckBox1[3]", "1");
                                    break;
                                case "Review at Next Operation":
                                    pdfField.SetField("CheckBox1[9]", "1");
                                    break;
                                case "Rework or Repair":
                                    pdfField.SetField("CheckBox1[10]", "1");
                                    break;
                                case "MRB Hold":
                                    pdfField.SetField("CheckBox2[0]", "1");
                                    break;
                            }
                            pdfField.SetField("TextField1[0]", qirObject.CurrentRevision.Problem);
                            pdfField.SetField("TextField2[0]", qirObject.PIC);
                            pdfField.SetField("TextField2[1]", qirObject.CurrentRevision.LotNumber);
                            pdfField.SetField("TextField2[2]", qirObject.CurrentRevision.RevSubmitter);
                            pdfField.SetField("TextField1[1]", qirObject.CurrentRevision.CauseReason);
                            pdfField.SetField("TextField2[3]", qirObject.WONumber.ToString());
                            pdfField.SetField("TextField2[4]", qirObject.CurrentRevision.SupplierID.ToString());
                            pdfField.SetField("TextField2[5]", qirObject.PartNumber);
                            pdfField.SetField("TextField2[6]", Convert.ToDateTime(qirObject.CurrentRevision.RevDate).ToShortTimeString());
                            pdfField.SetField("TextField2[7]", qirObject.CurrentRevision.NCMCode.ToString());
                            pdfField.SetField("TextField2[8]", qirObject.CurrentRevision.DiamondNumber);
                            pdfField.SetField("DateTimeField1[0]", qirObject.CurrentRevision.RevDate.ToShortDateString());
                            pdfField.SetField("TextField2[9]", qirObject.IDNumber.ToString());
                            pdfField.SetField("TextField2[10]", qirObject.CurrentRevision.MaterialLost.ToString());
                            pdfField.SetField("TextField1[2]", qirObject.CurrentRevision.DispositionReason);
                            pdfField.SetField("TextField2[11]", $"{qirObject.MaterialCost} / {qirObject.UOM}");
                            pdfField.SetField("TextField2[12]", Math.Round(Convert.ToDecimal(qirObject.MaterialCost * qirObject.CurrentRevision.MaterialLost), 2).ToString());
                            pdfField.SetField("TextField2[13]", qirObject.Found.ToString());
                            pdfField.SetField("TextField2[14]", qirObject.CurrentRevision.Origin.ToString());
                            pdfField.SetField("TextField3[0]", qirObject.CurrentRevision.Shift.ToString());
                            stamp.FormFlattening = false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
                }
                if (!autoSave)
                {
                    Process.Start($"{Properties.Settings.Default.omnitemp}{qirObject.IDNumber}.pdf");
                }
            }
        }

        /// <summary>
        /// Update an already loaded notice DataTable
        /// </summary>
        /// <param name="noticeTable">Current notice DataTable</param>
        public static void UpdateNoticeTable(this DataTable noticeTable)
        {
            using (DataTable _tempTable = QIR.LoadNotice(update:true))
            {
                if (_tempTable != null && _tempTable.Rows.Count > 0)
                {
                    foreach (DataRow row in _tempTable.Rows)
                    {
                        if (noticeTable.Select($"[QIRNumber] = '{row.ItemArray[2]}'").Count() == 0)
                        {
                            noticeTable.Rows.Add(row.ItemArray);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Check to see if the QIR object has photos attached to it
        /// </summary>
        /// <param name="qir">QIR Object</param>
        /// <returns>Photo(s) attached as boolean</returns>
        public static bool PhotoExists(this QIR qir)
        {
            return Directory.GetFiles(Properties.Settings.Default.QIRPhotoDirectory, $"{qir.IDNumber}P.*", SearchOption.TopDirectoryOnly).Length > 0;
        }

        /// <summary>
        /// Create a QIR Photo Micrsoft Word document and save selected photos to it
        /// </summary>
        /// <param name="qir">QIR Object</param>
        /// <param name="fileDrop">File Drop object used in the drag and drop function</param>
        /// <returns>Transaction Success as bool.  true = accepted / false = rejected</returns>
        public static bool AttachPhoto(this QIR qir, object fileDrop = null)
        {
            var _tempFile = new List<string>();
            if (fileDrop == null)
            {
                var ofd = new OpenFileDialog { Filter = "Camera Photos (*.jpg)|*.jpg|Cell Phone Photos (*.jpeg)|*.jpeg", Title = "Select the photo to add.", Multiselect = true };
                ofd.ShowDialog();
                if (ofd.FileNames.Length == 0)
                {
                    return false;
                }
                foreach (var s in ofd.FileNames)
                {
                    _tempFile.Add(s);
                }
            }
            else
            {
                foreach (var s in (string[])fileDrop)
                {
                    _tempFile.Add(s);
                }
            }
            var wordApp = new Word.Application
            {
                Visible = false
            };
            object empty = Missing.Value;
            var wordDoc = wordApp.Documents.Add(ref empty, ref empty, ref empty, ref empty);
            try
            {
                if (File.Exists($"{Properties.Settings.Default.QIRPhotoDirectory}{qir.IDNumber}P.docx"))
                {
                    File.SetAttributes($"{Properties.Settings.Default.QIRPhotoDirectory}{qir.IDNumber}P.docx", FileAttributes.Normal);
                    wordDoc = wordApp.Documents.Open($"{Properties.Settings.Default.QIRPhotoDirectory}{qir.IDNumber}P.docx");
                    foreach (var _file in _tempFile)
                    {
                        wordDoc.Application.Selection.InlineShapes.AddPicture(_file);
                    }
                }
                else
                {
                    foreach (Word.Section section in wordDoc.Sections)
                    {
                        var header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        header.Fields.Add(header, Word.WdFieldType.wdFieldPage);
                        header.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        header.Font.Size = 16;
                        header.Font.Bold = 5;
                        header.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                        header.Text = $"Quality Incident Report No. {qir.IDNumber} Photos";
                        var footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        footer.Fields.Add(footer, Word.WdFieldType.wdFieldPage);
                        footer.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        footer.Font.Size = 12;
                        footer.Text = $"Created by: {CurrentUser.FullName}              {DateTime.Now.ToLongDateString()}";
                        foreach (var _file in _tempFile)
                        {
                            wordDoc.Application.Selection.InlineShapes.AddPicture(_file);
                        }
                    }
                    wordDoc.SaveAs2($"{Properties.Settings.Default.QIRPhotoDirectory}{qir.IDNumber}P.docx");
                }
                wordDoc.Close(Word.WdSaveOptions.wdSaveChanges);
                wordDoc = null;
                wordApp.Quit(Word.WdSaveOptions.wdSaveChanges, ref empty, ref empty);
                wordApp = null;
                qir.IsPhotosAttached = true;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
                return false;
            }
            finally
            {
                if (wordDoc != null)
                {
                    wordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    wordDoc = null;
                }
                if (wordApp != null)
                {
                    wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                    wordApp = null;
                }
            }
        }

        /// <summary>
        /// Get QIR values from M2k ERP system
        /// </summary>
        /// <param name="qir">Current QIR Object</param>
        /// <param name="idNumber">Value to use as the M2k query</param>
        /// <param name="queryType">Query type</param>
        public static void GetQIRFromM2k(this QIR qir, string idNumber, M2kDataQuery queryType)
        {
            try
            {
                switch (queryType)
                {
                    case M2kDataQuery.LotNumber:
                        var _skew = new InventorySkew(idNumber);
                        if (_skew != null)
                        {
                            qir.PartNumber = _skew.PartNumber;
                            qir.CurrentRevision.DiamondNumber = _skew.DiamondNumber;
                            qir.MaterialCost = InventorySkew.GetMaterialCost(qir.PartNumber);
                            qir.UOM = _skew.UOM;
                            qir.WONumber = _skew.WorkOrderNumber;
                            qir.CurrentRevision.Origin = _skew.WorkCenter.IDNumber;
                            if (qir.QIRFormType == QIRType.QIREZ)
                            {
                                qir.Found = qir.CurrentRevision.Origin;
                                qir.NCMCodeList = null;
                            }
                            qir.Lost = null;
                        }
                        break;
                    case M2kDataQuery.WONumber:
                        var _qirData = InventorySkew.GetItemInformation(idNumber);
                        if (_qirData != null)
                        {
                            qir.PartNumber = _qirData[0];
                            qir.MaterialCost = InventorySkew.GetMaterialCost(qir.PartNumber);
                            qir.UOM = InventorySkew.GetUOM(qir.PartNumber);
                            qir.CurrentRevision.Origin = Convert.ToInt32(_qirData[1]);
                            if (string.IsNullOrEmpty(qir.CurrentRevision.LotNumber))
                            {
                                qir.CurrentRevision.LotNumber = "N/A";
                            }
                            if (qir.QIRFormType == QIRType.QIREZ)
                            {
                                qir.Found = qir.CurrentRevision.Origin;
                                qir.NCMCodeList = null;
                            }
                            qir.Lost = null;
                        }
                        break;
                    case M2kDataQuery.PartNumber:
                        qir.MaterialCost = InventorySkew.GetMaterialCost(idNumber);
                        qir.UOM = InventorySkew.GetUOM(idNumber);
                        qir.Lost = null;
                        break;
                }
            }
            catch (NullReferenceException)
            {
                return;
            }
        }

        /// <summary>
        /// Load the QIR List for parsing multiple QIR's at once
        /// </summary>
        /// <param name="lotNbr">QIR Object ID Number to load</param>
        public static void Load(this IList<QIR> _qirList, string lotNbr)
        {
            var _tempQIR = new List<int>();
            try
            {
                using (SqlCommand cmd = new SqlCommand($@"USE {App.DataBase};
                                                        SELECT [QIRNumber] FROM [qir_metrics_view] WHERE LotNumber=@p1", App.SqlConAsync))
                {
                    cmd.Parameters.AddWithValue("@p1", lotNbr);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                _tempQIR.Add(reader.SafeGetInt32("QIRNumber"));
                            }
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                foreach (int i in _tempQIR)
                {
                    _qirList.Add(new QIR(i, false));
                }
            }
            catch (Exception)
            { }
        }
    }
}
