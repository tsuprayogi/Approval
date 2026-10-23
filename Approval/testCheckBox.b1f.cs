using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using Application = SAPbouiCOM.Framework.Application;

namespace Approval
{
    [FormAttribute("Approval.testCheckBox", "testCheckBox.b1f")]
    class testCheckBox : UserFormBase
    {
        private SAPbouiCOM.Matrix Matrix0;

        private SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

        public testCheckBox()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_0").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_1").Specific));

            this.Matrix0.ChooseFromListAfter += this.Matrix0_ChooseFromListAfter;

            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            

        }

        private SAPbouiCOM.CheckBox CheckBox0;

        private void OnCustomInitialize()
        {
            SAPbouiCOM.IForm form = this.UIAPIRawForm;
            if (form == null)
            { throw new Exception("UIAPIRawForm belum tersedia."); }

            form.EnableMenu("1281", true);   //Find

            form.EnableMenu("1288", true);   // Next
            form.EnableMenu("1289", true);   // Prev
            form.EnableMenu("1290", true);   // First
            form.EnableMenu("1291", true);   // Last Record

            Matrix0.AddRow();

            Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
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

                //if (pVal.ItemUID == "Item_1" && pVal.ColUID == "colUser" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == true)
                //{
                //    // Catatan: Beberapa versi SDK membaca klik objek spesifik di dalam matrix sebagai et_ITEM_PRESSED
                //    Application.SBO_Application.MessageBox($"[SUKSES] Event ITEM_PRESSED terdeteksi pada Kolom Linked Button. Baris: {pVal.Row}");

                //    //BubbleEvent = false;
                //    return;
                //}

                //if (pVal.ItemUID == "Item_1" && pVal.ColUID == "colUser" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction == true)
                //{


                //    Application.SBO_Application.MessageBox($"[SUKSES] Event MATRIX_LINK_PRESSED terdeteksi pada Baris: {pVal.Row}, Kolom: {pVal.ColUID}");

                //    // Kita set false dulu agar SAP tidak mencoba membuka form default dan memicu error selama tes
                //    //BubbleEvent = false;
                //    return;
                //}

                //if (pVal.ItemUID == "Item_1" && pVal.ColUID == "colUser" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction == true)
                //{
                //    // 1. Ambil objek Form yang sedang aktif
                //    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                //    // 2. Ambil objek Matrix berdasarkan ItemUID ("Item_1")
                //    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_1").Specific;

                //    // 3. Ambil objek EditText/Cell dari baris (pVal.Row) dan kolom (pVal.ColUID) yang diklik
                //    SAPbouiCOM.EditText oCell = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colUser").Cells.Item(pVal.Row).Specific;

                //    // 4. Ambil value/nilai dari cell tersebut
                //    string cellValue = oCell.Value;

                //    // 5. Tampilkan nilai di MessageBox
                //    Application.SBO_Application.MessageBox($"[SUKSES] Baris: {pVal.Row}, Nilai colUser: {cellValue}");

                //    // Matikan default event jika Anda tidak ingin SAP B1 membuka form bawaan (seperti Master Data)
                //    BubbleEvent = false;
                //    return;
                //}

                if (pVal.ItemUID == "Item_1" && pVal.ColUID == "colUser" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction == true)
                {
                    // Pastikan kita membatalkan proses default SAP terlebih dahulu
                    BubbleEvent = false;

                    try
                    {
                        // 1. Ambil objek Form dan Matrix
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Item_1").Specific;

                        // 2. Ambil nilai U_NAME dari cell yang diklik
                        SAPbouiCOM.EditText oCell = (SAPbouiCOM.EditText)oMatrix.Columns.Item("colUser").Cells.Item(pVal.Row).Specific;
                        string cellValueName = oCell.Value.Trim();

                        if (string.IsNullOrEmpty(cellValueName))
                        {
                            return; // Jika cell kosong, tidak perlu memproses apapun
                        }

                        // 3. Ambil USERID dari database menggunakan Recordset (DI-API) berdasarkan U_NAME
                        // Catatan: Pastikan Anda sudah menginisialisasi objek Company (oCompany) dari Add-on Anda
                        SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        // Query untuk mencari USERID. Sesuaikan sintaks (" atau tanpa ") jika Anda menggunakan HANA / SQL.
                        string query = $"SELECT USERID FROM OUSR WHERE U_NAME = '{cellValueName}'";
                        oRecordset.DoQuery(query);

                        if (!oRecordset.EoF)
                        {
                            string userId = oRecordset.Fields.Item("USERID").Value.ToString();

                            // 4. Buka Form User (Object Type = 12) secara terprogram menggunakan USERID sebagai key data
                            // Parameter: OpenForm(BoFormObjectTypes, ObjectKeyStr, WindowTitle)
                            // Ganti SAPbouiCOM menjadi SAPbobsCOM
                           // Application.SBO_Application.OpenForm((SAPbobsCOM.BoFormObjectTypes)12, userId, "");

                            Application.SBO_Application.OpenForm(BoFormObjectEnum.fo_User,"12", userId);

                        }
                        else
                        {
                            Application.SBO_Application.SetStatusBarMessage($"User dengan nama '{cellValueName}' tidak ditemukan di sistem.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                DataTable selectedObjects = cflEvent.SelectedObjects;

                if (selectedObjects == null || selectedObjects.IsEmpty)
                    return;

                string userCode = Convert.ToString(
                    selectedObjects.GetValue("U_NAME", 0)
                ).Trim();

                

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


                if (string.IsNullOrEmpty(userCode))
                    return;

                int dsRow = pVal.Row - 1;

                SAPbouiCOM.DBDataSource dbDS = this.UIAPIRawForm.DataSources.DBDataSources.Item("@WTM1");

                dbDS.SetValue("U_U_UserID", dsRow, userCode);
                dbDS.SetValue("U_U_Dept", dsRow, deptName); // Mengisi Nama User
               // dbDS.SetValue("U_Department", dsRow, department); // Mengisi Department (Sesuaikan nama UDF Anda)


                this.Matrix0.SetLineData(pVal.Row);



                // Jika user mengisi data di BARIS TERAKHIR matrix, maka buatkan baris baru di bawahnya
                if (pVal.Row == this.Matrix0.RowCount)
                {
                    // Tambah baris kosong di DBDataSource
                    //dbDS.InsertRecord(dbDS.Size);

                    // Tambah baris kosong di UI Matrix (mengambil cetakan baris terakhir)
                    this.Matrix0.AddRow(1, this.Matrix0.RowCount);

                    this.Matrix0.FlushToDataSource();

                    // Opsional: Bersihkan baris baru di Matrix agar benar-benar kosong
                    this.Matrix0.ClearRowData(this.Matrix0.RowCount);
                }


            }
            catch (Exception ex)
            {
               Application.SBO_Application.MessageBox("Error saat memilih User: " + ex.Message);
            }
        }

        

    }
}
