
using System;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using SAPbobsCOM;

using Application = SAPbouiCOM.Framework.Application;
using System.Collections.Generic;

namespace Approval
{
    [FormAttribute("Approval.frmApprovalTemplate", "frmApprovalTemplate.b1f")]
    class frmApprovalTemplate : UserFormBase
    {
        //DataSource
        private DBDataSource oDSHeader;
        private DBDataSource oDSLines;
        private DBDataSource oDSStage;

        private SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();


        //Navigasi
        private string[] _keys = new string[0];
        private int _index = -1;
        private string _currentTemplateCode = "";

        // Status form dan navigasi
        private bool _isClosing = false;
        private bool _isInitialized = false;
        private string _formUID = string.Empty;

        //table field 
        private const string FIELD_USER_ID = "U_U_UserID";
        private const string COL_USER = "userCol";


        // =========================================================
        // COMPONENTS
        // =========================================================
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.CheckBox CheckBox0;

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;

        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Matrix Matrix1;

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private ComboBox ComboBox0;

        private StaticText StaticText4;
        private EditText EditText2;
        private StaticText StaticText5;


        public frmApprovalTemplate()
        {
        }

        private void SetPosition()
        {
            UIAPIRawForm.Left = 370;
            UIAPIRawForm.Top = 140;
        }

        // =========================================================
        // INITIALIZE COMPONENT
        // =========================================================
        public override void OnInitializeComponent()
        {
            this.SetPosition();
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_3").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_6").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_7").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_8").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_11").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_16").Specific));
            this.Matrix0.ChooseFromListAfter += this.Matrix0_ChooseFromListAfter;
            this.Matrix1.ChooseFromListAfter += this.Matrix1_ChooseFromListAfter;
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_2").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.OnCustomInitialize();

        }

        // =========================================================
        // FORM EVENTS
        // =========================================================
        public override void OnInitializeFormEvents()
        {
            //this.LoadAfter +=
            //    new LoadAfterHandler(this.Form_LoadAfter);
        }


        private void OnCustomInitialize()
        {
            try
            {
                SAPbouiCOM.IForm form = UIAPIRawForm;

                if (form == null)
                    throw new Exception("UIAPIRawForm belum tersedia.");

                _formUID = form.UniqueID;
                _isClosing = false;
                _isInitialized = false;

                // Navigasi SAP
                form.EnableMenu("1281", false); // Find
                form.EnableMenu("1288", true); // Next
                form.EnableMenu("1289", true); // Previous
                form.EnableMenu("1290", true); // First
                form.EnableMenu("1291", true); // Last

                //form.EnableMenu("1287", true); //duplicate
                //form.EnableMenu("1292", true); //add row
                form.EnableMenu("1293", true); //delete row
                form.EnableMenu("1294", true); //duplicate row
                form.EnableMenu("1283", true);

                LoadDataToComboBox(ComboBox0);

                // Matrix user
                if (Matrix0 == null)
                {
                    Matrix0 = (SAPbouiCOM.Matrix)
                        form.Items.Item("Item_8").Specific;
                }

                // Matrix stage
                if (Matrix1 == null)
                {
                    Matrix1 = (SAPbouiCOM.Matrix)
                        form.Items.Item("Item_11").Specific;
                }

                // Header dan detail datasource
                oDSHeader = form.DataSources.DBDataSources.Item("@OWTM");
                oDSLines = form.DataSources.DBDataSources.Item("@WTM1");
                oDSStage = form.DataSources.DBDataSources.Item("@WTM2");

                // Muat data menggunakan prosedur yang sudah ada.
                // Jangan membersihkan datasource sebelum pemuatan.
                LoadMatrixData();

                _currentTemplateCode = "";
                SetSaveButtonMode();

                _isInitialized = true;

                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;

                Application.SBO_Application.MenuEvent +=
       SBO_Application_MenuEvent;

            }
            catch (Exception ex)
            {
                _isInitialized = false;

                Application.SBO_Application.StatusBar.SetText(
                    "Error saat Initialize Approval Template:\n\n" +
                    ex.ToString(),BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning
                );
            }
        }



        private void SBO_Application_ItemEvent(
           string FormUID,
           ref SAPbouiCOM.ItemEvent pVal,
           out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (FormUID != UIAPIRawForm.UniqueID)
                    return;
          
        
                if (pVal.ItemUID == "Item_11" && pVal.ColUID == "colStage" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction == true)
                {
                    // Pastikan kita membatalkan proses default SAP terlebih dahulu
                    BubbleEvent = false;

                    try
                    {
                        // 1. Ambil objek Form dan Matrix
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_11").Specific;

                        // 2. Ambil nilai U_NAME dari cell yang diklik
                        SAPbouiCOM.EditText oCell = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colStage").Cells.Item(pVal.Row).Specific;
                        string cellValueName = oCell.Value.Trim();

                        if (string.IsNullOrEmpty(cellValueName))
                        {
                            return; // Jika cell kosong, tidak perlu memproses apapun
                        }

                        // 3. Ambil USERID dari database menggunakan Recordset (DI-API) berdasarkan U_NAME
                        // Catatan: Pastikan Anda sudah menginisialisasi objek Company (oCompany) dari Add-on Anda
                        SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        // Query untuk mencari USERID. Sesuaikan sintaks (" atau tanpa ") jika Anda menggunakan HANA / SQL.
                        string query = $"select  WstCode, Name from OWST  WHERE Name = '{cellValueName}'";
                        oRecordset.DoQuery(query);

                        if (!oRecordset.EoF)
                        {
                            string userId = oRecordset.Fields.Item("WstCode").Value.ToString();

                            // 4. Buka Form User (Object Type = 12) secara terprogram menggunakan USERID sebagai key data
                            // Parameter: OpenForm(BoFormObjectTypes, ObjectKeyStr, WindowTitle)
                            // Ganti SAPbouiCOM menjadi SAPbobsCOM
                            // Application.SBO_Application.OpenForm((SAPbobsCOM.BoFormObjectTypes)12, userId, "");

                            Application.SBO_Application.OpenForm(BoFormObjectEnum.fo_ConfirmationLevel, "120", userId);

                        }
                        else
                        {
                            Application.SBO_Application.SetStatusBarMessage($"Stage dengan nama '{cellValueName}' tidak ditemukan di sistem.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        }

                        // Bebaskan objek Recordset dari memori
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                    }
                    catch (Exception ex)
                    {
                        Application.SBO_Application.MessageBox($"Error saat membuka form User: {ex.Message}");
                    }

                    return;
                }

               

            }

            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(
                    "Error Approval:\n" + ex.Message);
            }


        }



        private void SBO_Application_MenuEvent(
      ref SAPbouiCOM.MenuEvent pVal,
      out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.MenuUID != "1283" || !pVal.BeforeAction)
                    return;

                SAPbouiCOM.Form activeForm =
                    Application.SBO_Application.Forms.ActiveForm;

                if (activeForm == null ||
                    activeForm.UniqueID != UIAPIRawForm.UniqueID)
                    return;

                // Batalkan Remove bawaan SAP.
                BubbleEvent = false;

                RemoveCurrentTemplate();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error Remove Approval Template: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        #region CFL

        private void Matrix0_ChooseFromListAfter(
            object sboObject,
            SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.Row <= 0)
                    return;

                ISBOChooseFromListEventArg cflEvent =
                    (ISBOChooseFromListEventArg)pVal;

                SAPbouiCOM.DataTable selectedObjects =
                    cflEvent.SelectedObjects;

                if (selectedObjects == null || selectedObjects.IsEmpty)
                    return;

                string userName = Convert.ToString(
                    selectedObjects.GetValue("U_NAME", 0)
                ).Trim();

                //string deptCode = Convert.ToString(
                //    selectedObjects.GetValue("Department", 0)
                //).Trim();

                string deptCode = Convert.ToString(
                   selectedObjects.GetValue("Department", 0)
               ).Trim();
                string deptName = "";


                // 2. QUERY KE OUDP: Cari Nama Departemen berdasarkan ID / Code
                if (!string.IsNullOrEmpty(deptCode))
                {
                    // Inisialisasi Recordset lewat DI API
                    SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    // Query untuk mengambil OUDP.Name berdasarkan OUDP.Code
                    string query = $"SELECT \"Name\" FROM OUDP WHERE \"Code\" = '{deptCode}'";
                    oRecordset.DoQuery(query);

                    if (!oRecordset.EoF)
                    {
                        deptName = Convert.ToString(oRecordset.Fields.Item("Name").Value).Trim();
                    }

                    // Lepaskan object dari memory (Clean up COM object)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                }

                if (string.IsNullOrWhiteSpace(userName))
                    return;


                SAPbouiCOM.Matrix matrix = Matrix0;
                SAPbouiCOM.DBDataSource ds = oDSLines;

                int row = pVal.Row;
                int dsRow = row - 1;

                // Pastikan record tersedia.
                while (ds.Size <= dsRow)
                {
                    ds.InsertRecord(ds.Size);
                }

                // PENTING: arahkan offset ke record yang sesuai
                // dengan baris matrix yang sedang diproses.
                ds.Offset = dsRow;

                // Baca nilai lama dari record yang benar.
                string before = ds.GetValue(
                    "U_U_UserID", ds.Offset
                ).Trim();

                // Tulis nilai user baru.
                ds.SetValue(
                    "U_U_UserID", ds.Offset, userName
                );

                ds.SetValue(
                    "U_U_Dept", ds.Offset, deptName
                );

                // Verifikasi setelah penulisan.
                string afterWrite = ds.GetValue(
                    "U_U_UserID", ds.Offset
                ).Trim();

                // Pindahkan nilai record yang benar ke baris matrix.
                matrix.SetLineData(row);

                // Periksa nilai yang tampil pada matrix.
                SAPbouiCOM.EditText userCell =
                    (SAPbouiCOM.EditText)matrix
                        .Columns.Item("userCol")
                        .Cells.Item(row).Specific;

                string matrixUser = userCell.Value.Trim();

                // Baca kembali nilai datasource.
                string afterSetLine = ds.GetValue(
                    "U_U_UserID", ds.Offset
                ).Trim();

                

                if (row == matrix.RowCount)
                {
                    matrix.AddRow(1, matrix.RowCount);
                    matrix.ClearRowData(matrix.RowCount);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error ChooseFromList: " + ex.Message,BoMessageTime.bmt_Medium,BoStatusBarMessageType.smt_Warning
                );
            }
        }


        private void Matrix1_ChooseFromListAfter(
            object sboObject,
            SBOItemEventArg pVal)
        {
            SAPbobsCOM.Recordset oRecordset = null;

            try
            {
                if (pVal.Row <= 0)
                    return;

                ISBOChooseFromListEventArg cflEvent =
                    (ISBOChooseFromListEventArg)pVal;

                SAPbouiCOM.DataTable selectedObjects =
                    cflEvent.SelectedObjects;

                if (selectedObjects == null ||
                    selectedObjects.IsEmpty)
                {
                    return;
                }

                // 1. Ambil data Approval Template dari CFL
                string templateCode = Convert.ToString(
                    selectedObjects.GetValue("WstCode", 0)
                ).Trim();

                string remarks = Convert.ToString(
                    selectedObjects.GetValue("Remarks", 0)
                ).Trim();

                if (string.IsNullOrEmpty(templateCode))
                    return;

                string stageName = "";
               // string approverIDs = "";

                // 2. Query nama stage dari OWST
                oRecordset = (SAPbobsCOM.Recordset)
                    oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string safeCode = templateCode.Replace("'", "''");

                string queryStage = $@"
            SELECT ""Name""
            FROM OWST
            WHERE ""WstCode"" = '{safeCode}'";

                oRecordset.DoQuery(queryStage);

                if (!oRecordset.EoF)
                {
                    stageName = Convert.ToString(
                        oRecordset.Fields.Item("Name").Value
                    ).Trim();
                }

                // 3. Ambil semua UserID dari WST1
                string approverNames = "";

                string queryApproverNames = $@"
                    SELECT STRING_AGG(
                        CAST(U.""U_NAME"" AS NVARCHAR(100)), ', '
                    ) AS ""U_U_Approver""
                    FROM WST1 W
                    INNER JOIN OUSR U
                        ON U.""USERID"" = W.""UserID""
                    WHERE W.""WstCode"" = '{safeCode}'";

                oRecordset.DoQuery(queryApproverNames);

                if (!oRecordset.EoF)
                {
                    object value =
                        oRecordset.Fields.Item("U_U_Approver").Value;

                    if (value != null && value != DBNull.Value)
                        approverNames = Convert.ToString(value).Trim();
                }

             
                // 4. Lepaskan Recordset
                System.Runtime.InteropServices.Marshal
                    .ReleaseComObject(oRecordset);

                oRecordset = null;

                // 5. Simpan ke DBDataSource @WTM2
                SAPbouiCOM.Matrix matrix1 = Matrix1;
                SAPbouiCOM.DBDataSource ds = oDSStage;

                int row = pVal.Row;
                int dsRow = row - 1;

                while (ds.Size <= dsRow)
                {
                    ds.InsertRecord(ds.Size);
                }

                ds.Offset = dsRow;

                // Simpan nama stage
                ds.SetValue(
                    "U_U_StageCode",
                    ds.Offset,
                    stageName);

                // Simpan remarks
                ds.SetValue(
                    "U_U_Remarks",
                    ds.Offset,
                    remarks);

                // Simpan semua UserID dengan separator koma
                ds.SetValue(
                    "U_U_Approver",
                    ds.Offset,
                    approverNames);
                
                // Refresh tampilan matrix
                matrix1.SetLineData(row);
                              
                // Tambahkan baris baru jika memilih baris terakhir
                if (row == matrix1.RowCount)
                {
                    matrix1.AddRow(1, matrix1.RowCount);
                    matrix1.ClearRowData(matrix1.RowCount);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Error saat memilih Approval Template:\n\n" +
                    ex.ToString(),BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
                if (oRecordset != null)
                {
                    System.Runtime.InteropServices.Marshal
                        .ReleaseComObject(oRecordset);
                }
            }
        }

      
        
        #endregion

        #region Matrix

        // LOAD MATRIX DATA

        private void LoadMatrixData()
        {
            try
            {
                if (Matrix0 == null)
                    throw new Exception("Matrix0 belum tersedia. Item_8.");

                if (Matrix1 == null)
                    throw new Exception("Matrix1 belum tersedia. Item_11.");

                SAPbouiCOM.IForm form = UIAPIRawForm;

                oDSHeader = form.DataSources.DBDataSources.Item("@OWTM");
                oDSLines = form.DataSources.DBDataSources.Item("@WTM1");
                oDSStage = form.DataSources.DBDataSources.Item("@WTM2");

                // Bersihkan tampilan matrix.
                // Jangan menghapus datasource di sini jika form
                // digunakan juga untuk membuka template tersimpan.
                Matrix0.Clear();
                Matrix1.Clear();

                // Siapkan satu baris input kosong untuk user.
                Matrix0.AddRow();
                Matrix0.ClearRowData(1);

                // Siapkan satu baris input kosong untuk stage.
                Matrix1.AddRow();
                Matrix1.ClearRowData(1);
            }
            catch (Exception ex)
            {
                throw new Exception(
                    "LoadMatrixData Error:\n\n" + ex.ToString());
            }
        }

        private void EnsureDataSourceRow(int row)
        {
            if (oDSLines == null)
            {
                oDSLines =
                    this.UIAPIRawForm.DataSources.DBDataSources
                        .Item("@WTM1");
            }

            while (oDSLines.Size <= row)
            {
                oDSLines.InsertRecord(oDSLines.Size);
            }
        }

        private void EnsureStageSourceRow(int row)
        {
            if (oDSStage == null)
            {
                oDSStage =
                    this.UIAPIRawForm.DataSources.DBDataSources
                        .Item("@WTM2");
            }

            while (oDSStage.Size <= row)
            {
                oDSStage.InsertRecord(oDSStage.Size);
            }
        }

        private void SetRowNumber(SAPbouiCOM.Matrix matrix)
        {
            try
            {
                if (matrix == null)
                    return;

                for (int i = 1; i <= matrix.RowCount; i++)
                {
                    try
                    {
                        SAPbouiCOM.EditText txt =
                            (SAPbouiCOM.EditText)
                            matrix.Columns.Item("#")
                            .Cells.Item(i)
                            .Specific;

                        txt.Value = i.ToString();
                    }
                    catch
                    {
                        // Jangan hentikan proses hanya karena
                        // nomor row gagal diisi.
                    }
                }
            }
            catch
            {
                // Ignore
            }
        }

        private string GetMatrixCellValue(SAPbouiCOM.Matrix matrix, int row,
   string columnId)
        {
            try
            {
                if (matrix == null)
                    return "";

                if (row <= 0 || row > matrix.VisualRowCount)
                    return "";

                SAPbouiCOM.Cell oCell =
                    matrix.Columns.Item(columnId)
                          .Cells.Item(row);

                object specific = oCell.Specific;

                if (specific is SAPbouiCOM.EditText editText)
                    return editText.Value.Trim();

                if (specific is SAPbouiCOM.ComboBox comboBox)
                    return comboBox.Value.Trim();

                return "";
            }
            catch
            {
                return "";
            }
        }


        #endregion

        
        #region Navigation

        public void LoadKeys()
        {
            try
            {
                string q = @"
            SELECT Code
            FROM [@OWTM]
            ORDER BY Code";

                SAPbobsCOM.Recordset rs =
                    (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                rs.DoQuery(q);

                List<string> list = new List<string>();

                while (!rs.EoF)
                {
                    string code =
                        Convert.ToString(
                            rs.Fields.Item("Code").Value).Trim();

                    if (!string.IsNullOrEmpty(code))
                        list.Add(code);

                    rs.MoveNext();
                }

                _keys = list.ToArray();

                // Belum ada record aktif
                _index = -1;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "LoadKeys gagal:\n\n" + ex.ToString(),BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
            }
        }


        // ==========================================================
        // FIRST
        // ==========================================================

        public string First()
        {
            if (_keys == null || _keys.Length == 0)
                return null;

            _index = 0;

            return _keys[_index];
        }


        // ==========================================================
        // LAST
        // ==========================================================

        public string Last()
        {
            if (_keys == null || _keys.Length == 0)
                return null;

            _index = _keys.Length - 1;

            return _keys[_index];
        }


        // ==========================================================
        // NEXT
        // ==========================================================

        public string Next()
        {
            if (_keys == null || _keys.Length == 0)
                return null;

            if (_index < 0)
            {
                _index = 0;
                return _keys[_index];
            }

            if (_index < _keys.Length - 1)
                _index++;

            return _keys[_index];
        }


        // ==========================================================
        // PREVIOUS
        // ==========================================================

        public string Previous()
        {
            if (_keys == null || _keys.Length == 0)
                return null;

            if (_index < 0)
            {
                _index = _keys.Length - 1;
                return _keys[_index];
            }

            if (_index > 0)
                _index--;

            return _keys[_index];
        }
        
        private void SyncCurrentIndex()
        {
            try
            {
                if (_keys == null || _keys.Length == 0)
                    return;

                // Karena Code header kita tidak ditampilkan di EditText,
                // kita ambil Code berdasarkan data template yang sedang aktif.

                string currentCode = GetCurrentTemplateCode();

                if (string.IsNullOrWhiteSpace(currentCode))
                    return;

                for (int i = 0; i < _keys.Length; i++)
                {
                    if (string.Equals(
                        _keys[i],
                        currentCode,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        _index = i;
                        return;
                    }
                }
            }
            catch
            {
            }
        }
        
        public void HandleNavigation(string menuUID)
        {
            try
            {
                // Abaikan jika form sudah ditutup atau belum selesai initialize.
                if (_isClosing || !_isInitialized)
                    return;

                // Pastikan form masih bisa diakses.
                SAPbouiCOM.IForm form;

                try
                {
                    form = UIAPIRawForm;

                    if (form == null)
                        return;

                    // Pastikan instance form masih sama.
                    if (string.IsNullOrEmpty(_formUID) ||
                        form.UniqueID != _formUID)
                    {
                        return;
                    }
                }
                catch
                {
                    _isClosing = true;
                    _isInitialized = false;
                    return;
                }

                // Pastikan Company tersedia.
                if (oCompany == null || !oCompany.Connected)
                {
                    Application.SBO_Application.MessageBox(
                        "Company SAP belum terkoneksi.");
                    return;
                }

               

                // ADD
                if (menuUID == "1282")
                {
                    PrepareAddMode();
                    return;
                }

                // Hanya proses menu navigasi yang dikenal.
                if (menuUID != "1288" &&
                    menuUID != "1289" &&
                    menuUID != "1290" &&
                    menuUID != "1291")
                {
                    return;
                }

                if (_keys == null || _keys.Length == 0)
                {
                    LoadKeys();
                }

                if (_isClosing)
                    return;

                if (_keys == null || _keys.Length == 0)
                {
                    Application.SBO_Application.MessageBox(
                        "Belum ada Approval Template.");
                    return;
                }

                string code = null;

                switch (menuUID)
                {
                    case "1288":
                        code = Next();
                        break;

                    case "1289":
                        code = Previous();
                        break;

                    case "1290":
                        code = First();
                        break;

                    case "1291":
                        code = Last();
                        break;
                }

                if (_isClosing)
                    return;

                if (!string.IsNullOrWhiteSpace(code))
                {
                    LoadApprovalTemplate(code);
                }
            }
            catch (Exception ex)
            {
                if (!_isClosing)
                {
                    Application.SBO_Application.MessageBox(
                        "Handle Navigation gagal:\n\n" +
                        ex.ToString());
                }
            }
        }
        
        #endregion

        #region AppLoad
        private string GetCurrentTemplateCode()
        {
            return _currentTemplateCode;
        }

        private void LoadApprovalTemplate(string templateCode)
        {
            SAPbouiCOM.IForm oForm = null;
            bool isFrozen = false;

            // Ambil object form aktif terlebih dahulu
           // SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;

            try
            {
                if (_isClosing || !_isInitialized)
                    return;
               

                if (string.IsNullOrWhiteSpace(templateCode))
                    return;

                try
                {
                    oForm = UIAPIRawForm;

                    if (oForm == null)
                        return;

                    if (string.IsNullOrEmpty(_formUID) ||
                        oForm.UniqueID != _formUID)
                    {
                        return;
                    }
                }
                catch
                {
                    return;
                }

                // 1. BEKUKAN FORM (Mencegah UI glitch dan mengoptimalkan rendering)
                oForm.Freeze(true);
                isFrozen = true;

                SAPbobsCOM.Recordset rsHeader = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string qHeader = @"
            SELECT Code, Name, U_U_Active, U_TName, U_TRemarks, U_ObjectType, U_TCondition
            FROM [@OWTM]
            WHERE Code = '" + templateCode.Replace("'", "''") + @"'";

                rsHeader.DoQuery(qHeader);

                if (rsHeader.EoF)
                {
                    // Ingat untuk mencairkan form jika keluar di tengah jalan
                    oForm.Freeze(false);
                    Application.SBO_Application.MessageBox("Approval Template tidak ditemukan:\n" + templateCode);
                    return;
                }

                // Header mapping
                EditText0.Value = Convert.ToString(rsHeader.Fields.Item("U_TName").Value).Trim();
                EditText1.Value = Convert.ToString(rsHeader.Fields.Item("U_TRemarks").Value).Trim();

                string active = Convert.ToString(rsHeader.Fields.Item("U_U_Active").Value).Trim();
                CheckBox0.Checked = active.Equals("Y", StringComparison.OrdinalIgnoreCase);

                string objectType = Convert.ToString(rsHeader.Fields.Item("U_ObjectType").Value).Trim();
                if (!string.IsNullOrWhiteSpace(objectType))
                {
                    try { ComboBox0.Select(objectType, SAPbouiCOM.BoSearchKey.psk_ByValue); } catch { }
                }

                EditText2.Value = Convert.ToString(rsHeader.Fields.Item("U_TCondition").Value).Trim();

                LoadApprovalUsers(templateCode);
                LoadApprovalStages(templateCode);

                _currentTemplateCode = templateCode;

                SetSaveButtonMode();

                if (_keys != null)
                {
                    for (int i = 0; i < _keys.Length; i++)
                    {
                        if (string.Equals(_keys[i], templateCode, StringComparison.OrdinalIgnoreCase))
                        {
                            _index = i;
                            break;
                        }
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsHeader);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Load Approval Template gagal:\n\n" + ex.Message);
            }
            finally
            {
                //// 2. SELALU CAIRKAN FORM DI SINI (Memaksa semua item UI aktif dan redraw)
                //if (oForm != null)
                //{
                //    oForm.Freeze(false);
                //    oForm.Update(); // Paksa windows update grafis komponen
                //}

                if (isFrozen && oForm != null)
                {
                    try
                    {
                        oForm.Freeze(false);
                    }
                    catch
                    {
                        // Form mungkin sudah ditutup. 
                    }
                }
            }
        }


 
        private void LoadApprovalUsers(string templateCode)
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                if (oDSLines == null)
                {
                    oDSLines = UIAPIRawForm.DataSources.DBDataSources
                        .Item("@WTM1");
                }

                // Bersihkan data record sebelumnya
                Matrix0.Clear();
                oDSLines.Clear();

                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                    SAPbobsCOM.BoObjectTypes.BoRecordset);

                string safeCode = templateCode.Replace("'", "''");

                string query = @"
                    SELECT
                        T0.U_U_UserID,
                        T0.U_U_Dept AS DeptCode,
                        T1.Name AS DeptName
                    FROM [@WTM1] T0
                    LEFT JOIN OUDP T1
                        ON T0.U_U_Dept = T1.Code
                    WHERE T0.Code LIKE '" + safeCode + @"-%'
                    ORDER BY T0.Code";

                rs.DoQuery(query);

                int row = 0;

                while (!rs.EoF)
                {
                    row++;

                    // Tambah satu baris, lalu isi baris yang baru dibuat
                    Matrix0.AddRow();

                    string userID = Convert.ToString(
                        rs.Fields.Item("U_U_UserID").Value).Trim();

                    string deptName = Convert.ToString(
                        rs.Fields.Item("DeptName").Value).Trim();

                    string deptCode = Convert.ToString(
                        rs.Fields.Item("DeptCode").Value).Trim();

                    Matrix0.SetCellWithoutValidation(
                        row, COL_USER, userID);

                    Matrix0.SetCellWithoutValidation(
                        row, "deptCol", deptName);

                    // Sinkronkan datasource untuk setiap baris
                    //oDSLines.InsertRecord(oDSLines.Size - 0 > row - 1
                    //    ? row - 1
                    //    : oDSLines.Size);

                            int dsRow = row - 1;

                            while (oDSLines.Size <= dsRow)
                            {
                                oDSLines.InsertRecord(oDSLines.Size);
                            }

                            oDSLines.SetValue("U_U_UserID", dsRow, userID);
                            oDSLines.SetValue("U_U_Dept", dsRow, deptCode);

                            rs.MoveNext();
                }

                // Baris kosong untuk input user baru
                Matrix0.AddRow();

                int blankRow = Matrix0.VisualRowCount;
                Matrix0.SetCellWithoutValidation(blankRow, COL_USER, "");
                Matrix0.SetCellWithoutValidation(blankRow, "deptCol", "");
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(
                    "Load Approval Users gagal:\n\n" + ex.ToString());
            }
            finally
            {
                if (rs != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }    
        
        private void LoadApprovalStages(string templateCode)
        {
            SAPbobsCOM.Recordset rs = null;

            try
            {
                if (oDSStage == null)
                {
                    oDSStage = UIAPIRawForm.DataSources.DBDataSources
                        .Item("@WTM2");
                }

                Matrix1.Clear();
                oDSStage.Clear();

                 rs =
                    (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string safeCode = templateCode.Replace("'", "''");


                string q = @"
    SELECT
        T1.Name,
        T0.U_U_SortId,
        T0.U_U_Remarks,
        ISNULL(AP.ApproverNames, '') AS U_U_Approver
    FROM [@WTM2] T0
    LEFT JOIN [OWST] T1
        ON T0.U_U_StageCode = T1.WstCode

    OUTER APPLY
    (
        SELECT STRING_AGG(
            CAST(U.U_NAME AS NVARCHAR(MAX)), ', '
        ) AS ApproverNames
        FROM STRING_SPLIT(T0.U_U_Approver, ',') S
        INNER JOIN OUSR U
            ON U.USERID = TRY_CAST(
                LTRIM(RTRIM(S.value)) AS INT
            )
    ) AP

    WHERE T0.Code LIKE '" + safeCode + @"-%'
    ORDER BY T0.U_U_SortId";


                rs.DoQuery(q);

                int row = 0;

                while (!rs.EoF)
                {
                    row++;

                    Matrix1.AddRow();



                    string stageId = Convert.ToString(
                        rs.Fields.Item("Name").Value).Trim();

                    string remark = Convert.ToString(
                        rs.Fields.Item("U_U_Remarks").Value).Trim();

                    string approverNames = Convert.ToString(
                        rs.Fields.Item("U_U_Approver").Value).Trim();

                    Matrix1.SetCellWithoutValidation(
                        row, "colStage", stageId);

                    Matrix1.SetCellWithoutValidation(
                        row, "colDesc", remark);

                    // Tampilkan nama approver di kolom matrix
                    Matrix1.SetCellWithoutValidation(
                        row, "colApprove", approverNames);


                    int dsRow = row - 1;

                    while (oDSStage.Size <= dsRow)
                    {
                        oDSStage.InsertRecord(oDSStage.Size);
                    }

                    oDSStage.SetValue("Name", dsRow, stageId);
                    oDSStage.SetValue("U_U_Remarks", dsRow, remark);
                    oDSStage.SetValue("U_U_Approver", dsRow, approverNames);

                    rs.MoveNext();
                }
               
                // Baris kosong untuk input user baru
                Matrix1.AddRow();

                int blankRow = Matrix1.VisualRowCount;
                Matrix1.SetCellWithoutValidation(blankRow, "colStage", "");
                Matrix1.SetCellWithoutValidation(blankRow, "colDesc", "");
                Matrix1.SetCellWithoutValidation(blankRow, "colApprove", "");

            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(
                    "Load Approval Stages gagal:\n\n" +
                    ex.Message);
            }
        }


        


        private void PrepareAddMode()
        {
            SAPbouiCOM.IForm form = UIAPIRawForm;
            bool frozen = false;

            try
            {
                form.Freeze(true);
                frozen = true;

                _currentTemplateCode = "";
                _index = -1;

                // Pastikan datasource tersedia
                if (oDSHeader == null)
                    oDSHeader = form.DataSources.DBDataSources.Item("@OWTM");

                if (oDSLines == null)
                    oDSLines = form.DataSources.DBDataSources.Item("@WTM1");

                if (oDSStage == null)
                    oDSStage = form.DataSources.DBDataSources.Item("@WTM2");

                // Bersihkan datasource
                oDSHeader.Clear();
                oDSLines.Clear();
                oDSStage.Clear();

                // Siapkan record header untuk Add Mode
                oDSHeader.InsertRecord(0);

                // Masuk ke Add Mode secara eksplisit
                form.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                // Kosongkan header
                EditText0.Value = "";
                EditText1.Value = "";
                EditText2.Value = "";

                // Set nilai Active melalui datasource
                oDSHeader.SetValue("U_U_Active", 0, "Y");
                CheckBox0.Checked = false;

                // Reset ComboBox
                try
                {
                    if (ComboBox0.ValidValues.Count > 0)
                    {
                        ComboBox0.Select(
                            0,
                            SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                }
                catch
                {
                    // Tidak ada nilai default yang tersedia
                }

                // Reset matrix
                Matrix0.Clear();
                Matrix1.Clear();

                Matrix0.LoadFromDataSource();
                Matrix1.LoadFromDataSource();

                // Sediakan satu baris input kosong pada setiap matrix
                Matrix0.AddRow();
                Matrix0.ClearRowData(1);

                Matrix1.AddRow();
                Matrix1.ClearRowData(1);

                SetSaveButtonMode();

                form.EnableMenu("1282", true);

                form.Update();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Prepare Add Mode gagal: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (frozen)
                {
                    try
                    {
                        form.Freeze(false);
                    }
                    catch
                    {
                        // Form mungkin sudah ditutup
                    }
                }
            }
        }


        #endregion

        private void SaveApprovalTemplate()
        {
            SAPbobsCOM.Company oCompany = null;

            try
            {
                // =========================================================
                // 1. VALIDASI COMPANY
                // =========================================================

                oCompany = Program.oCompany;

                if (oCompany == null || !oCompany.Connected)
                {
                    throw new Exception(
                        "SAP Company belum terkoneksi.");
                }

                // =========================================================
                // 2. AMBIL DATA HEADER
                // =========================================================

                string templateName =
                    EditText0.Value.Trim();

                string templateRemarks =
                    EditText1.Value.Trim();

                string objectType =
                    ComboBox0.Value.Trim();

                string active =
                    CheckBox0.Checked ? "Y" : "N";

                string condition = EditText2.Value.Trim();

                // =========================================================
                // 3. VALIDASI HEADER
                // =========================================================

                if (string.IsNullOrEmpty(templateName))
                {
                    Application.SBO_Application.MessageBox(
                        "Template Name wajib diisi.");

                    EditText0.Active = true;
                    return;
                }

                if (string.IsNullOrEmpty(objectType))
                {
                    Application.SBO_Application.MessageBox(
                        "Object Type wajib dipilih.");

                    ComboBox0.Active = true;
                    return;
                }

                // =========================================================
                // 4. FLUSH MATRIX
                // =========================================================

                try
                {
                    Matrix0.FlushToDataSource();
                }
                catch
                {
                    // Abaikan
                }

                try
                {
                    Matrix1.FlushToDataSource();
                }
                catch
                {
                    // Abaikan
                }

                // =========================================================
                // 5. GENERATE CODE
                // =========================================================

                bool isUpdateMode =
                    !string.IsNullOrWhiteSpace(_currentTemplateCode);

                string templateCode;

                if (isUpdateMode)
                {
                    // Pertahankan Code record lama
                    templateCode = _currentTemplateCode;
                }
                else
                {
                    // Generate Code hanya saat Add
                    templateCode = "APT-" +
                        DateTime.Now.ToString("yyyyMMddHHmmss");
                }

                
                // =========================================================
                // 6. SIMPAN HEADER @OWTM
                // =========================================================

                SAPbobsCOM.UserTable oHeader = null;
                int resultHeader;

                try
                {
                    oHeader = oCompany.UserTables.Item("OWTM");

                    if (isUpdateMode)
                    {
                        // Ambil record yang akan di-update terlebih dahulu.
                        if (!oHeader.GetByKey(templateCode))
                        {
                            throw new Exception(
                                "Approval Template tidak ditemukan: " +
                                templateCode);
                        }
                    }
                    else
                    {
                        // Siapkan record baru.
                        oHeader.Code = templateCode;
                        oHeader.Name = templateCode;
                    }

                    // Isi field SETELAH GetByKey() pada mode Update.
                    oHeader.UserFields.Fields.Item("U_U_Active").Value =
                        active;

                    oHeader.UserFields.Fields.Item("U_TName").Value =
                        templateName;

                    oHeader.UserFields.Fields.Item("U_TRemarks").Value =
                        templateRemarks;

                    oHeader.UserFields.Fields.Item("U_ObjectType").Value =
                        objectType;

                    oHeader.UserFields.Fields.Item("U_TCondition").Value =
                        condition;

                    // Simpan header.
                    if (isUpdateMode)
                    {
                        resultHeader = oHeader.Update();
                    }
                    else
                    {
                        resultHeader = oHeader.Add();
                    }

                    if (resultHeader != 0)
                    {
                        int errorCode;
                        string errorMessage;

                        oCompany.GetLastError(
                            out errorCode,
                            out errorMessage);

                        throw new Exception(
                            (isUpdateMode
                                ? "Gagal Update @OWTM."
                                : "Gagal Add @OWTM.") +
                            "\n\n" +
                            errorCode + " - " + errorMessage);
                    }

                    if (isUpdateMode)
                    {
                        DeleteTemplateDetails(
                            oCompany, "WTM1", templateCode);

                        DeleteTemplateDetails(
                            oCompany, "WTM2", templateCode);
                    }
                    
                }
                finally
                {
                    if (oHeader != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(
                            oHeader);
                    }
                }


                // =========================================================
                // 7. SIMPAN MATRIX0 -> @WTM1
                // =========================================================

                int userLine = 1;

                // 1. Inisialisasi UserTable di luar loop agar lebih efisien
                SAPbobsCOM.UserTable oUser = oCompany.UserTables.Item("WTM1");
                // 2. Siapkan objek Recordset untuk convert Name ke Code
                SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                try
                {
                    for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                    {
                        string userID = GetMatrixCellValue(Matrix0, i, COL_USER);
                        string deptName = GetMatrixCellValue(Matrix0, i, "deptCol"); // Ini berisi nama (misal: "BOD")

                        // -----------------------------------------------------
                        // Abaikan row kosong
                        // -----------------------------------------------------
                        if (string.IsNullOrWhiteSpace(userID))
                            continue;

                        string deptCode = ""; // Variable untuk menampung Code hasil convert

                        // -----------------------------------------------------
                        // CONVERT NAMA DEPT KE CODE (OUDP)
                        // -----------------------------------------------------
                        if (!string.IsNullOrWhiteSpace(deptName))
                        {
                            string query = $"SELECT \"Code\" FROM OUDP WHERE \"Name\" = '{deptName.Replace("'", "''")}'";
                            oRecordset.DoQuery(query);

                            if (!oRecordset.EoF)
                            {
                                deptCode = Convert.ToString(oRecordset.Fields.Item("Code").Value).Trim();
                            }
                        }

                        // -----------------------------------------------------
                        // Code harus unik
                        // -----------------------------------------------------
                        string detailCode = templateCode + "-" + userLine.ToString("000");

                        oUser.Code = detailCode;
                        oUser.Name = detailCode;

                        oUser.UserFields.Fields.Item("U_U_UserID").Value = userID;

                        // Simpan nilainya menggunakan DEPT CODE, bukan nama teks
                        oUser.UserFields.Fields.Item("U_U_Dept").Value = deptCode;

                        int resultUser = oUser.Add();

                        if (resultUser != 0)
                        {
                            int errorCode;
                            string errorMessage;

                            oCompany.GetLastError(
                                out errorCode,
                                out errorMessage);

                            throw new Exception(
                                "Gagal menyimpan @WTM1.\n" +
                                "Baris Matrix       : " + i + "\n" +
                                "Detail Code        : " + detailCode + "\n" +
                                "Template Code      : " + templateCode + "\n" +
                                "User ID            : " + userID + "\n" +
                                "Department Code    : " + deptCode + "\n" +
                                "SAP Error Code     : " + errorCode + "\n" +
                                "SAP Error Message  : " + errorMessage);
                        }

                        userLine++;
                    }
                }
                finally
                {
                    // 3. Wajib bersihkan COM Object memory setelah loop selesai
                    if (oRecordset != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                    }
                    if (oUser != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUser);
                    }
                }



               
                // =========================================================
                // 8. SIMPAN MATRIX1 -> @WTM2
                // =========================================================

                int stageLine = 1;

                // 1. Inisialisasi UserTable di luar loop agar lebih efisien
                SAPbobsCOM.UserTable oStage = oCompany.UserTables.Item("WTM2");
                // 2. Siapkan objek Recordset untuk convert Name ke Code
                SAPbobsCOM.Recordset oReStage = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                try
                {
                    for (int i = 1;
                     i <= Matrix1.VisualRowCount;
                     i++)
                    {
                        string stageCode =
                            GetMatrixCellValue(
                                Matrix1,
                                i,
                                "colStage");

                        string remarks =
                            GetMatrixCellValue(
                                Matrix1,
                                i,
                                "colDesc");

                        string approves =
                            GetMatrixCellValue(
                                Matrix1,
                                i,
                                "colApprove");                       

                        // Konversi nama approver menjadi USERID
                        string approverIDs = "";

                        if (!string.IsNullOrWhiteSpace(approves))
                        {
                            string[] approverNames = approves.Split(',');

                            List<string> ids = new List<string>();

                            foreach (string approverName in approverNames)
                            {
                                string name = approverName.Trim();

                                if (string.IsNullOrWhiteSpace(name))
                                    continue;

                                string safeName = name.Replace("'", "''");

                                string queryApprover = $@"
                                    SELECT ""USERID""
                                    FROM OUSR
                                    WHERE ""U_NAME"" = '{safeName}'";

                                oReStage.DoQuery(queryApprover);

                                if (!oReStage.EoF)
                                {
                                    string userId = Convert.ToString(
                                        oReStage.Fields.Item("USERID").Value).Trim();

                                    if (!string.IsNullOrWhiteSpace(userId))
                                        ids.Add(userId);
                                }
                                else
                                {
                                    throw new Exception(
                                        "Approver '" + name +
                                        "' tidak ditemukan di tabel OUSR.");
                                }
                            }

                            approverIDs = string.Join(",", ids);
                        }



                        string stageId = ""; // Variable untuk menampung Code hasil convert

                        // -----------------------------------------------------
                        // CONVERT NAMA DEPT KE CODE (OUDP)
                        // -----------------------------------------------------
                        if (!string.IsNullOrWhiteSpace(stageCode))
                        {
                            string query = $"SELECT \"WstCode\" FROM OWST WHERE \"Name\" = '{stageCode.Replace("'", "''")}'";
                            oReStage.DoQuery(query);

                            if (!oReStage.EoF)
                            {
                                stageId = Convert.ToString(oReStage.Fields.Item("WstCode").Value).Trim();
                            }
                        }


                       
                        // -----------------------------------------------------
                        // Abaikan row kosong
                        // -----------------------------------------------------

                        if (string.IsNullOrWhiteSpace(stageCode))
                            continue;

                        //SAPbobsCOM.UserTable oStage =
                        //    oCompany.UserTables.Item("WTM2");

                        // -----------------------------------------------------
                        // Code harus unik
                        // -----------------------------------------------------

                        string detailCode =
                            templateCode + "-" +
                            stageLine.ToString("000");

                        oStage.Code = detailCode;
                        oStage.Name = detailCode;

                        oStage.UserFields.Fields.Item("U_U_StageCode").Value = stageId;

                        oStage.UserFields.Fields.Item("U_U_SortId").Value = i.ToString();

                        oStage.UserFields.Fields.Item("U_U_Remarks").Value = remarks;

                        oStage.UserFields.Fields.Item("U_U_Approver").Value = approverIDs;



                        int resultStage =
                            oStage.Add();

                        if (resultStage != 0)
                        {
                            throw new Exception(
                                "Gagal menyimpan @WTM2 pada baris " +
                                i +
                                ".\n\n" +
                                oCompany.GetLastErrorDescription());
                        }

                        stageLine++;
                    }
                }
                finally
                {
                    // 3. Wajib bersihkan COM Object memory setelah loop selesai
                    if (oReStage != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oReStage);
                    }
                    if (oStage != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oStage);
                    }
                }



                // =========================================================
                // 9. SUCCESS
                // =========================================================


                Application.SBO_Application.StatusBar.SetText(
                    "Approval Template berhasil " +
                    (isUpdateMode ? "di-update." : "ditambahkan.") +
                    " | Code: " + templateCode +
                    " | Name: " + templateName,
                    SAPbouiCOM.BoMessageTime.bmt_Medium,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);


                // Refresh daftar navigasi
                LoadKeys();

                // Tetap buka template yang baru disimpan
                _currentTemplateCode = templateCode;
                LoadApprovalTemplate(templateCode);

                SetSaveButtonMode();
                
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "Save Approval Template gagal:\n\n" +
                    ex.ToString(), BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
            }
        }

        public void MarkFormClosing()
        {
            _isClosing = true;
            _isInitialized = false;
        }



        private void DeleteTemplateDetails(
            SAPbobsCOM.Company company,
            string tableName,
            string templateCode)
        {
            SAPbobsCOM.Recordset rs = null;
            SAPbobsCOM.UserTable detail = null;

            try
            {
                string safeCode = templateCode.Replace("'", "''");

                rs = (SAPbobsCOM.Recordset)
                    company.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Ambil seluruh Code detail milik template ini.
                // Hindari menghapus detail milik template lain.
                string query = $@"
            SELECT Code
            FROM [@{tableName}]
            WHERE Code LIKE '{safeCode}-%'
            ORDER BY Code";

                rs.DoQuery(query);

                List<string> detailCodes = new List<string>();

                while (!rs.EoF)
                {
                    string code = Convert.ToString(
                        rs.Fields.Item("Code").Value).Trim();

                    if (!string.IsNullOrEmpty(code))
                        detailCodes.Add(code);

                    rs.MoveNext();
                }

                detail = company.UserTables.Item(tableName);

                foreach (string code in detailCodes)
                {
                    if (!detail.GetByKey(code))
                    {
                        throw new Exception(
                            $"Detail {tableName} tidak ditemukan: {code}");
                    }

                    int result = detail.Remove();

                    if (result != 0)
                    {
                        throw new Exception(
                            $"Gagal menghapus {tableName}, Code {code}: " +
                            company.GetLastErrorDescription());
                    }
                }
            }
            finally
            {
                if (detail != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(
                        detail);

                if (rs != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }



        public void RemoveCurrentTemplate()
        {
            SAPbobsCOM.Company company = null;
            SAPbobsCOM.UserTable header = null;

            string templateCode = _currentTemplateCode;

            try
            {
                // 1. Validasi template yang sedang dibuka
                if (string.IsNullOrWhiteSpace(templateCode))
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Pilih Approval Template yang ingin dihapus terlebih dahulu.",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    return;
                }

                company = Program.oCompany;

                if (company == null || !company.Connected)
                    throw new Exception("SAP Company belum terkoneksi.");

                // 2. Konfirmasi sebelum menghapus
                int answer = Application.SBO_Application.MessageBox(
                    "Apakah Anda yakin ingin menghapus Approval Template ini?\n\n" +
                    "Template Code : " + templateCode + "\n" +
                    "Template Name : " + EditText0.Value + "\n\n" +
                    "Data header dan seluruh detail template akan dihapus.",
                    2,
                    "Ya",
                    "Tidak");

                // Tombol pertama adalah Ya.
                if (answer != 1)
                    return;

                // 3. Mulai transaksi
                company.StartTransaction();

                // 4. Hapus detail user terlebih dahulu
                DeleteTemplateDetails(
                    company,
                    "WTM1",
                    templateCode);

                // 5. Hapus detail stage
                DeleteTemplateDetails(
                    company,
                    "WTM2",
                    templateCode);

                // 6. Hapus header template
                header = company.UserTables.Item("OWTM");

                if (!header.GetByKey(templateCode))
                {
                    throw new Exception(
                        "Approval Template tidak ditemukan: " +
                        templateCode);
                }

                int result = header.Remove();

                if (result != 0)
                {
                    int errorCode;
                    string errorMessage;

                    company.GetLastError(
                        out errorCode,
                        out errorMessage);

                    throw new Exception(
                        "Gagal menghapus header @OWTM.\n" +
                        errorCode + " - " + errorMessage);
                }

                // 7. Commit transaksi
                if (company.InTransaction)
                {
                    company.EndTransaction(
                        SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }

                // 8. Refresh navigasi
                _currentTemplateCode = "";
                _index = -1;

                LoadKeys();

                // 9. Kembali ke form Add Mode
                PrepareAddMode();

                Application.SBO_Application.StatusBar.SetText(
                    "Approval Template berhasil dihapus: " +
                    templateCode,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                // Batalkan transaksi jika masih berjalan.
                try
                {
                    if (company != null && company.InTransaction)
                    {
                        company.EndTransaction(
                            SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                }
                catch
                {
                    // Pertahankan pesan error utama.
                }

                Application.SBO_Application.StatusBar.SetText(
                    "Remove Approval Template gagal: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (header != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(
                        header);
                }
            }
        }


        #region Button
        private void Button0_ClickBefore(
        object sboObject,
        SBOItemEventArg pVal,
        out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SaveApprovalTemplate();
            }
            catch (Exception ex)
            {
                BubbleEvent = false;

                Application.SBO_Application.StatusBar.SetText(
                    "Error Save Approval Template:\n\n" +
                    ex.ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
            }
        }


        public void LoadDataToComboBox(SAPbouiCOM.ComboBox oComboBox)
        {
            SAPbobsCOM.Recordset oRecordset = null;

            SAPbouiCOM.IForm oForm = UIAPIRawForm;

            try
            {
                // Blokir redrawing form agar proses loading lebih cepat dan tidak berkedip
                oForm.Freeze(true);

                // Bersihkan data ComboBox yang lama jika ada
                while (oComboBox.ValidValues.Count > 0)
                {
                    oComboBox.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                // Inisialisasi Recordset dari DI API
                oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Tulis query SQL (Contoh: mengambil data dari tabel User-Defined Table '@MY_TABLE')
                // U_Code sebagai Value (Hidden), U_Name sebagai Description (Teks yang muncul)
                string query = "SELECT U_T2_ObjCode, U_T2_ObjName FROM [@SBO_OBJECT] ORDER BY U_T2_ObjCode";
                oRecordset.DoQuery(query);

                // Loop hasil query dan masukkan ke ComboBox
                while (!oRecordset.EoF)
                {
                    string code = oRecordset.Fields.Item("U_T2_ObjCode").Value.ToString();
                    string name = oRecordset.Fields.Item("U_T2_ObjName").Value.ToString();

                    oComboBox.ValidValues.Add(code, name);
                    oRecordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Error: " + ex.Message);
            }
            finally
            {
                // Lepas objek Recordset dari memori
                if (oRecordset != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                    oRecordset = null;
                }
                // Buka kembali pembekuan form
                oForm.Freeze(false);
            }
        }


        private void Button1_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            UIAPIRawForm.Close();
        }


        private void SetSaveButtonMode()
        {
            if (Button0 == null)
                return;

            bool isUpdateMode =
                !string.IsNullOrWhiteSpace(_currentTemplateCode);

            Button0.Caption = isUpdateMode ? "Update" : "Add";
        }

        #endregion

        
    }
}
