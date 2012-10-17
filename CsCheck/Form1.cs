using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;

namespace CsCheck
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //app.config編碼
            //MyClass.ToggleConfigEncryption("CsCheck.exe");
        }
        //設定程序用ie開啟連結
        public System.Diagnostics.Process p = new System.Diagnostics.Process();
        //常數設定
        Encoding BIG5 = Encoding.GetEncoding("big5");
        bool m_IsPortOpened = false; //讀卡機通訊埠狀態
        bool m_IsSAMChecked = false; //讀卡機SAM卡認證狀態
        bool m_IsHPCChecked = false; //醫是人員卡認證狀態
        int nErrCode; //錯誤代碼
        Form2 f2 = new Form2(); //訊息視窗


        //功能設定
        [DllImport("CsHis.dll", EntryPoint = "csOpenCom")] //開啟讀卡機
        private static extern int csOpenCom(int pComNum);
        [DllImport("CsHis.dll", EntryPoint = "csCloseCom")]//關閉讀卡機
        private static extern int csCloseCom();
        [DllImport("CsHis.dll", EntryPoint = "csVerifySAMDC")]//與健保局連線
        private static extern int csVerifySAMDC();
        [DllImport("CsHis.dll", EntryPoint = "hpcVerifyHPCPIN")]//檢查醫事人員卡之PIN值
        private static extern int hpcVerifyHPCPIN();
        [DllImport("CsHis.dll", EntryPoint = "hisGetBasicData")]//讀取個人資料
        private static extern int hisGetBasicData(byte[] pBuffer, ref int iBufferLen);
        //ERRORCODE hisReadPrescription(char *pOutpatientPrescription， int *iBufferLenOutpatient，char *pLongTermPrescription， int *iBufferLenLongTerm，char *pImportantTreatmentCode， int *iBufferLenImportant，char *pIrritationDrug， int *iBufferLenIrritation);
        [DllImport("CsHis.dll", EntryPoint = "hisReadPrescription")]//讀取處方箋作業
        private static extern int hisReadPrescription(byte[] pOutpatientPrescription, ref int iBufferLenOutpatient, byte[] pLongTermPrescription, ref int iBufferLenLongTerm, byte[] pImportantTreatmentCode, ref int iBufferLenImportant, byte[] pIrritationDrug, ref int iBufferLenIrritation);
        [DllImport("CsHis.dll")]//取得控制軟體版本
        private static extern int csGetVersionEx(byte[] pPath);
        [DllImport("CsHis.dll")]//讀取卡片狀態
        private static extern int hisGetCardStatus(int CardType);
        [DllImport("CsHis.dll")]//讀取就醫資料不需HPC卡的部分
        private static extern int hisGetTreatmentNoNeedHPC(byte[] pBuffer, ref int iBufferLen);
        [DllImport("CsHis.dll")]//讀取就醫資料需HPC卡的部分
        private static extern int hisGetTreatmentNeedHPC(byte[] pBuffer, ref int iBufferLen);
        //保險對象特定醫療資訊查詢作業
        [DllImport("PEAT7403B01.dll")]
        private static extern int PEA_SamExeNhiQuery(byte[] sHostName, int nPort,byte[] sBusCode,int nCom,byte[] sHcaId,byte[] sPatId,byte[] sPatBirth);
        //讀取錯誤訊息
        [DllImport("PEAT7403B01.dll")]
        private static extern void PEA_GetMsg(byte[] sBuf,ref int nSize);
        //讀取醫事人員卡身分證
        [DllImport("CsHis.dll")]
        private static extern int hpcGetHPCSSN(byte[] SSN, ref int Len_SSN);


        //richtextbox控制游標停在最下面
        void downTextbox()
        {
            rtOutput.SelectionStart = rtOutput.Text.Length;
            rtOutput.ScrollToCaret();
        }

        //檢查讀卡機認證狀態
        bool checkSAMStatus()
        {
            openCom();
            int status = hisGetCardStatus(1);
            if (status == 2)
            {
                label3.Text = "讀卡機認證狀態：已認證";
                m_IsSAMChecked = true;
            }
            else
            {
                label3.Text = "讀卡機認證狀態：未認證";
                m_IsSAMChecked = false;
            }
            closeCom();
            return m_IsSAMChecked;
        }

        //寫入訊息
        void msg(string msg)
        {
            
            string text = DateTime.Now.ToLongTimeString() + "：" + msg + "\r\n";
            rtOutput.Text += text;
            downTextbox();
        }

        //檢查醫事人員卡認證狀態
        bool checkHPCStatus()
        {
            openCom();
            int status = hisGetCardStatus(3);
            if (status == 3)
            {
                label4.Text = "醫事卡認證狀態：已認證";
                m_IsHPCChecked = true;
            }
            else
            {
                label4.Text = "醫事卡認證狀態：未認證";
                m_IsHPCChecked = false;
            }
            closeCom();
            return m_IsHPCChecked;
        }

        //檢查健保卡卡狀態
        bool checkCardStatus()
        {
            openCom();
            int status = hisGetCardStatus(2);
            closeCom();
            switch (status)
            {
                case 0:
                    MessageBox.Show("卡片未置入");
                    return false;

                case 1:
                    MessageBox.Show("健保IC卡尚未與安全模組認證");
                    return false;
                case 9:
                    MessageBox.Show("所置入非健保IC卡");
                    return false;
                default:
                    return true;
            }
        }

        //讀卡機連線
        void openCom()
        {
            if ((nErrCode = csOpenCom(0)) == 0)
            {
                m_IsPortOpened = true;
                return;
            }
            else
            {
                msg("開啟讀卡機通訊埠，失敗！請檢查讀卡機連線。");
                return;
            }
        }

        //讀卡機斷線
        void closeCom()
        {
            nErrCode = csCloseCom();
            if (nErrCode == 0)
            {
                m_IsPortOpened = false;
                return;
            }
            else
            {
                return;
            }
        }

        //SAM認證
        void verifySAM()
        {       
            //if (checkSAMStatus() == true)
            //{
            //    msg("讀卡機認證成功。");
            //    m_IsSAMChecked = true;
            //    label3.Text = "讀卡機認證狀態：已認證";
            //    return;
            //}
            //openCom();
            Form2 f2 = new Form2();
            f2.showMessage("與健保局連線認證中，請稍後...");
            f2.Show();
            Application.DoEvents();

            this.Cursor = Cursors.WaitCursor;

            nErrCode = csVerifySAMDC();

            f2.Close();
            f2.Dispose();

            this.Cursor = Cursors.Default;
            if (nErrCode == 0)
            {
                msg("讀卡機認證成功。");
                m_IsSAMChecked = true;
                label3.Text = "讀卡機認證狀態：已認證";
            }
            else
            {
                msg("讀卡機認證失敗" + "\r\n" + "錯誤訊息：" + ErrCode.errMsg(nErrCode));
            }
            //closeCom();
        }

        //醫事人員卡認證
        void verifyHPC()
        {
            //openCom();
            f2.showMessage("與健保局連線認證中\r\n當出現密碼輸入視窗時\r\n請於讀卡機上輸入醫事人員卡密碼...");
            f2.Show();
            Application.DoEvents();

            this.Cursor = Cursors.WaitCursor;

            nErrCode = hpcVerifyHPCPIN();

            f2.Hide();
            this.Cursor = Cursors.Default;

            if (nErrCode == 0)
            {
                msg("醫事卡認證成功。");
                label4.Text = "醫事卡認證狀態：已認證 ";
                m_IsHPCChecked = true;
            }
            else
            {
                msg("醫事卡認證失敗" + "\r\n" + "錯誤原因：" + ErrCode.errMsg(nErrCode));
            }
            //closeCom();
        }

        //讀取個人資料
        void getBasicData()
        {
            openCom();
            int buff = 72;
            byte[] pBuffer = new byte[buff];
            nErrCode = hisGetBasicData(pBuffer, ref buff);
            closeCom();
            if (nErrCode != 0)
            {
                msg("讀取錯誤。" + ErrCode.errMsg(nErrCode));
                return;
            }
            else
            {
                Encoding BIG5 = Encoding.GetEncoding("big5");
                string CardNo = BIG5.GetString(pBuffer, 0, 12).Trim();
                string Name = BIG5.GetString(pBuffer, 12, 20).Trim();
                string PID = BIG5.GetString(pBuffer, 32, 10).Trim();
                string Birthday = BIG5.GetString(pBuffer, 42, 7).Trim();
                string Gender = BIG5.GetString(pBuffer, 49, 1).Trim();
                string DeliverDate = BIG5.GetString(pBuffer, 50, 7).Trim();
                string VoidFlag = BIG5.GetString(pBuffer, 57, 1).Trim();
                string EmergencyPhoneNumber = BIG5.GetString(pBuffer, 58, 14).Trim();

                label1.Text = "卡片號碼：" + CardNo;
                label5.Text = "姓名：" + Name;
                label6.Text = "身分證號：" + PID;
                label7.Text = "出生日期：" + Birthday;
                label8.Text = "性別：" + Gender;
                label9.Text = "發卡日期：" + DeliverDate;
                label10.Text = "卡片註銷註記：" + VoidFlag;
                label11.Text = "緊急聯絡電話：" + EmergencyPhoneNumber;
            }
            msg("讀取卡片個人資料成功。");
        }
        //dataGridView 變色，警示藥物
        private void ChangeColor(DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count - 1; i++)
            {
                if (dgv.Rows[i].Cells[5].Value.ToString().Contains("ZOLPIDEM"))
                {
                        dgv.Rows[i].Cells[3].Style.BackColor = Color.Yellow;
                        dgv.Rows[i].Cells[4].Style.BackColor = Color.Yellow;
                        dgv.Rows[i].Cells[5].Style.BackColor = Color.Yellow;
                        //dgv.Rows[i].Cells[3].Value = "kkkkkk";
                }
            }
            for (int i = 0; i < dgv.Rows.Count - 1; i++)
            {
                if (dgv.Rows[i].Cells[5].Value.ToString().Contains("FLUNITRAZEPAM"))
                {
                    dgv.Rows[i].Cells[3].Style.BackColor = Color.Red;
                    dgv.Rows[i].Cells[4].Style.BackColor = Color.Red;
                    dgv.Rows[i].Cells[5].Style.BackColor = Color.Red;
                }
            }

        }
        //讀取處方資料
        void hisReadPrescription()
        {
            openCom();
            //門診處方箋
            int iBufferLenOutpatient = 3660;
            byte[] pOutpatientPrescription = new byte[iBufferLenOutpatient];
            //長期處方箋
            int iBufferLenLongTerm = 1320;
            byte[] pLongTermPrescription = new byte[iBufferLenLongTerm];
            //重要醫令
            int iBufferLenImportant = 360;
            byte[] pImportantTreatmentCode = new byte[iBufferLenImportant];
            //過敏藥物
            int iBufferLenIrritation = 120;
            byte[] pIrritationDrug = new byte[iBufferLenIrritation];

            f2.showMessage("讀取卡片「處方籤」中\r\n依資料長度需15～30秒\r\n請稍後...");
            f2.Show();
            Application.DoEvents();

            this.Cursor = Cursors.WaitCursor;

            nErrCode = hisReadPrescription(pOutpatientPrescription, ref iBufferLenOutpatient, pLongTermPrescription, ref iBufferLenLongTerm, pImportantTreatmentCode, ref iBufferLenImportant, pIrritationDrug, ref iBufferLenIrritation);

            f2.Hide();
            this.Cursor = Cursors.Default;

            if (nErrCode != 0)
            {
                msg("讀取錯誤。" + ErrCode.errMsg(nErrCode));
                closeCom();
                return;
            }
            else
            {
                DataColumn column;
                DataRow row;
                //門診處方箋
                DataTable dt1 = new DataTable();
                column = new DataColumn();
                column.ColumnName = "序號";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "就診日期";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "醫令類別";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "代號";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "藥品名稱或項目代號";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "藥品成分";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "診療部位";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "用法";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "天數";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "總量";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "交付處方註記";
                dt1.Columns.Add(column);
                dataGridView1.DataSource = dt1;
                
                //迴圈讀取處方籤
                for (int i = 0; i < 60; i++)
                {
                    //刪除數字前的0
                    //string reg = @"[1-9]\d+";

                    row = dt1.NewRow();
                    row[0] = i + 1;//序號
                    //row[1] = BIG5.GetString(pOutpatientPrescription, (i * 61) + 0, 13).Trim().Substring(0, 7);//日期
                    string date = BIG5.GetString(pOutpatientPrescription, (i * 61) + 0, 13).Trim().Substring(0, 7);
                    row[1] = date;
                    row[2] = BIG5.GetString(pOutpatientPrescription, (i * 61) + 13, 1).Trim();//醫令類別
                    row[3] = BIG5.GetString(pOutpatientPrescription, (i * 61) + 14, 12).Trim();//藥品或診療項目代號
                    string id = BIG5.GetString(pOutpatientPrescription, (i * 61) + 14, 12).Trim();
                    string sqlString = "SELECT [name] FROM [GHHP].[dbo].[druglibrary] where id = '" + id + "' and startdate < '" + date + "' and enddate > '" + date + "'";
                    string result = MyClass.ExecuteScalar(sqlString);
                    if (result == "0")
                    {
                        row[4] = id;
                    }
                    else
                    {
                        row[4] = result;
                    }
                    //藥品成分
                    string sqlString2 = "SELECT [ingredient] FROM [GHHP].[dbo].[druglibrary] where id = '" + id + "' and startdate < '" + date + "' and enddate > '" + date + "'";
                    string result2 = MyClass.ExecuteScalar(sqlString2);
                    if (result2 == "0")
                    {
                        row[5] = id;
                    }
                    else
                    {
                        row[5] = result2;
                    }

                    row[6] = BIG5.GetString(pOutpatientPrescription, (i * 61) + 26, 6).Trim();//診療部位
                    row[7] = BIG5.GetString(pOutpatientPrescription, (i * 61) + 32, 18).Trim();//用法
                    row[8] = BIG5.GetString(pOutpatientPrescription, (i * 61) + 50, 2).Trim();//天數
                    row[9] = Convert.ToDouble(BIG5.GetString(pOutpatientPrescription, (i * 61) + 52, 7).Trim()) / 100;//總量
                    //row[7] = System.Text.RegularExpressions.Regex.Match(BIG5.GetString(pOutpatientPrescription, (i * 61) + 52, 7).Trim(), reg);
                    row[10] = BIG5.GetString(pOutpatientPrescription, (i * 61) + 59, 2).Trim();//交付處方註記
                    dt1.Rows.Add(row);
                    
                }
                dataGridView1.Columns[0].Width = 40;
                dataGridView1.Columns[1].Width = 100;
                dataGridView1.Columns[2].Width = 40;
                dataGridView1.Columns[3].Width = 100;
                dataGridView1.Columns[4].Width = 300;
                dataGridView1.Columns[5].Width = 200;
                dataGridView1.Columns[6].Width = 50;
                dataGridView1.Columns[7].Width = 80;
                dataGridView1.Columns[8].Width = 50;
                dataGridView1.Columns[9].Width = 80;
                dataGridView1.Columns[10].Width = 50;
                ChangeColor(dataGridView1);

                //長期處方箋
                DataTable dt2 = new DataTable();
                column = new DataColumn();
                column.ColumnName = "序號";
                dt2.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "開立日期";
                dt2.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "藥品代碼";
                dt2.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "藥品名稱";
                dt2.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "用法";
                dt2.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "天數";
                dt2.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "總量";
                dt2.Columns.Add(column);
                dataGridView2.DataSource = dt2;

                //迴圈讀取資料
                for (int i = 0; i < 30; i++)
                {
                    row = dt2.NewRow();
                    row[0] = i + 1;//序號
                    string date = BIG5.GetString(pLongTermPrescription, (i * 44) + 0, 7).Trim();
                    row[1] = date;
                    row[2] = BIG5.GetString(pLongTermPrescription, (i * 44) + 7, 10).Trim();
                    string id = BIG5.GetString(pLongTermPrescription, (i * 44) + 7, 10).Trim();
                    string sqlString = "SELECT [name] FROM [GHHP].[dbo].[druglibrary] where id = '" + id + "' and startdate < '" + date + "' and enddate > '" + date + "'";
                    string result = MyClass.ExecuteScalar(sqlString);
                    if (result == "0")
                    {
                        row[3] = id;
                    }
                    else
                    {
                        row[3] = result;
                    }
                    row[4] = BIG5.GetString(pLongTermPrescription, (i * 44) + 17, 18).Trim();
                    row[5] = BIG5.GetString(pLongTermPrescription, (i * 44) + 35, 2).Trim();
                    row[6] = BIG5.GetString(pLongTermPrescription, (i * 44) + 37, 7).Trim();
                    dt2.Rows.Add(row);
                }

                dataGridView2.Columns[0].Width = 40;
                dataGridView2.Columns[1].Width = 100;
                dataGridView2.Columns[2].Width = 100;
                dataGridView2.Columns[3].Width = 300;
                dataGridView2.Columns[4].Width = 100;
                dataGridView2.Columns[5].Width = 100;
                dataGridView2.Columns[6].Width = 100;
                ChangeColor(dataGridView2);

                //重要醫令
                DataTable dt3 = new DataTable();
                column = new DataColumn();
                column.ColumnName = "序號";
                dt3.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "實施日期";
                dt3.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "醫療院所代碼";
                dt3.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "重要醫令項目代碼";
                dt3.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "實施部位代碼";
                dt3.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "總量";
                dt3.Columns.Add(column);
                dataGridView3.DataSource = dt3;

                //迴圈讀取資料
                for (int i = 0; i < 10; i++)
                {
                    row = dt3.NewRow();
                    row[0] = i + 1;//序號
                    string date = BIG5.GetString(pImportantTreatmentCode, (i * 36) + 0, 7).Trim();
                    row[1] = date;
                    row[2] = BIG5.GetString(pImportantTreatmentCode, (i * 36) + 7, 10).Trim();
                    row[3] = BIG5.GetString(pImportantTreatmentCode, (i * 36) + 17, 6).Trim();
                    row[4] = BIG5.GetString(pImportantTreatmentCode, (i * 36) + 23, 6).Trim();
                    row[5] = BIG5.GetString(pImportantTreatmentCode, (i * 36) + 29, 7).Trim();
                    dt3.Rows.Add(row);
                }

                dataGridView3.Columns[0].Width = 40;
                dataGridView3.Columns[1].Width = 100;
                dataGridView3.Columns[2].Width = 100;
                dataGridView3.Columns[3].Width = 300;
                dataGridView3.Columns[4].Width = 100;
                dataGridView3.Columns[5].Width = 100;
                ChangeColor(dataGridView3);

                //過敏藥物
                DataTable dt4 = new DataTable();
                column = new DataColumn();
                column.ColumnName = "序號";
                dt4.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "過敏藥物成份";
                dt4.Columns.Add(column);
                dataGridView4.DataSource = dt4;

                //迴圈讀取資料
                for (int i = 0; i < 3; i++)
                {
                    row = dt4.NewRow();
                    row[0] = i + 1;//序號
                    row[1] = BIG5.GetString(pIrritationDrug, (i * 40) + 0, 40).Trim();
                    dt4.Rows.Add(row);
                }

                dataGridView4.Columns[0].Width = 40;
                dataGridView4.Columns[1].Width = 500;

                //取消排序
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                for (int i = 0; i < dataGridView2.Columns.Count; i++)
                {
                    dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                for (int i = 0; i < dataGridView3.Columns.Count; i++)
                {
                    dataGridView3.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                for (int i = 0; i < dataGridView4.Columns.Count; i++)
                {
                    dataGridView4.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
            }

            msg("卡片「處方箋」讀取作業，完成。");
            closeCom();


        }

        //讀取就醫資料
        void hisGetTreatmentNoNeedHPC()
        {
            openCom();
            //就醫資料不需HPC的部份
            int iBufferLen1 = 498;
            byte[] pBuffer1 = new byte[iBufferLen1];
            int iBufferLen2 = 372;
            byte[] pBuffer2 = new byte[iBufferLen2];

            f2.showMessage("讀取卡片「就醫資料」中\r\n依資料長度需10～20秒\r\n請稍後...");
            f2.Show();
            Application.DoEvents();

            this.Cursor = Cursors.WaitCursor;

            nErrCode = hisGetTreatmentNoNeedHPC(pBuffer1, ref iBufferLen1);
            int nErrCode2 = hisGetTreatmentNeedHPC(pBuffer2, ref iBufferLen2);

            f2.Hide();
            this.Cursor = Cursors.Default;

            if (nErrCode != 0 )
            {
                msg("讀取錯誤。" + ErrCode.errMsg(nErrCode));
                closeCom();
                return;
            }
            else if (nErrCode2 != 0)
            {
                msg("讀取錯誤。" + ErrCode.errMsg(nErrCode2));
                closeCom();
                return;
            }
            else
            {
                DataColumn column;
                DataRow row;
                //就醫資料
                DataTable dt1 = new DataTable();
                column = new DataColumn();
                column.ColumnName = "序號";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "就醫類別";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "新生兒就醫註記";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "就診日期時間";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "補卡註記";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "就醫序號";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "醫療院所";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "主要診斷碼";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "次要診斷碼";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "門診醫療費用";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "門診部分負擔費用";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "住院醫療費用";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "住院部分負擔費用(急性30天、慢性180天以內)";
                dt1.Columns.Add(column);
                column = new DataColumn();
                column.ColumnName = "住院部分負擔費用(急性31天、慢性180天以上)";
                dt1.Columns.Add(column);
                dataGridView5.DataSource = dt1;

                //迴圈讀取資料
                for (int i = 0; i < 6; i++)
                {
                    row = dt1.NewRow();
                    row[0] = i + 1;//序號
                    row[1] = BIG5.GetString(pBuffer1, (i * 69) + 84, 2).Trim();
                    row[2] = BIG5.GetString(pBuffer1, (i * 69) + 86, 1).Trim();
                    //row[3] = BIG5.GetString(pBuffer1, (i * 69) + 87, 13).Trim();
                    row[3] = BIG5.GetString(pBuffer1, (i * 69) + 87, 3).Trim() + "/" + BIG5.GetString(pBuffer1, (i * 69) + 90, 2).Trim() + "/" + BIG5.GetString(pBuffer1, (i * 69) + 92, 2).Trim() + " , " + BIG5.GetString(pBuffer1, (i * 69) + 94, 2).Trim() + ":" + BIG5.GetString(pBuffer1, (i * 69) + 96, 2).Trim() + ":" + BIG5.GetString(pBuffer1, (i * 69) + 98, 2).Trim();
                    row[4] = BIG5.GetString(pBuffer1, (i * 69) + 100, 1).Trim();
                    row[5] = BIG5.GetString(pBuffer1, (i * 69) + 101, 4).Trim();
                    //醫療院所代碼
                    //row[6] = BIG5.GetString(pBuffer1, (i * 69) + 105, 10).Trim();
                    string id = BIG5.GetString(pBuffer1, (i * 69) + 105, 10).Trim();
                    string sqlString = "SELECT [hosp] FROM [GHHP].[dbo].[hosplib] WHERE id = '" + id + "'";
                    var result = MyClass.ExecuteScalar(sqlString);
                    if (result == "0")
                    {
                        row[6] = id;
                    }
                    else
                    {
                        row[6] = result;
                    }
                    //主診斷碼
                    row[7] = BIG5.GetString(pBuffer2, (i * 43) + 127, 5).Trim();
                    //次診斷碼
                    row[8] = BIG5.GetString(pBuffer2, (i * 43) + 132, 5).Trim() + " / " + BIG5.GetString(pBuffer2, (i * 43) + 137, 5).Trim() + " / " + BIG5.GetString(pBuffer2, (i * 43) + 142, 5).Trim() + " / " + BIG5.GetString(pBuffer2, (i * 43) + 147, 5).Trim() + " / " + BIG5.GetString(pBuffer2, (i * 43) + 152, 5).Trim();
                    row[9] = BIG5.GetString(pBuffer1, (i * 69) + 115, 8).Trim();
                    row[10] = BIG5.GetString(pBuffer1, (i * 69) + 123, 8).Trim();
                    row[11] = BIG5.GetString(pBuffer1, (i * 69) + 131, 8).Trim();
                    row[12] = BIG5.GetString(pBuffer1, (i * 69) + 139, 7).Trim();
                    row[13] = BIG5.GetString(pBuffer1, (i * 69) + 146, 7).Trim();
                    dt1.Rows.Add(row);
                }

                dataGridView5.Columns[0].Width = 40;
                dataGridView5.Columns[1].Width = 40;
                dataGridView5.Columns[2].Width = 100;
                dataGridView5.Columns[3].Width = 150;
                dataGridView5.Columns[4].Width = 40;
                dataGridView5.Columns[5].Width = 60;
                dataGridView5.Columns[6].Width = 100;
                dataGridView5.Columns[7].Width = 100;
                dataGridView5.Columns[8].Width = 150;
                dataGridView5.Columns[9].Width = 100;
                dataGridView5.Columns[10].Width = 100;
                dataGridView5.Columns[11].Width = 100;
                dataGridView5.Columns[10].Width = 100;
                dataGridView5.Columns[11].Width = 100;
            }

            for (int i = 0; i < dataGridView5.Columns.Count; i++)
            {
                dataGridView5.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            msg("卡片「就醫資料」讀取作業，完成。");
            closeCom();
        }

        //
        void PEA_SamExeNhiQuery()
        {
            if (m_IsSAMChecked == false)
            {
                msg("請先進行讀卡機認證。");
                return;
            }
            if (m_IsHPCChecked == false)
            {
                msg("請先進行醫事卡認證。");
                return;
            }
            //private static extern int PEA_SamExeNhiQuery(byte[] sHostName, int nPort,byte[] sBusCode,int nCom,byte[] sHcaId,byte[] sPatId,byte[] sPatBirth);
            openCom();
            //健保局vpn ip
            string ip = "10.253.253.242";
            byte[] sHostName = new byte[14];
            sHostName = System.Text.Encoding.Default.GetBytes(ip);
            //健保局vpn port
            int nPort = 7403;
            //目前提供「功能類別」為保險對象特定醫療資訊查詢作業，請填入01。
            byte[] sBusCode = new byte[2];
            string busCode = "01";
            sBusCode = System.Text.Encoding.Default.GetBytes(busCode);
            //讀卡機port
            int nCom = 0;
            //private static extern int hpcGetHPCSSN(byte[] SSN, ref int Len_SSN);醫事人員卡身分證
            int Len_SSN = 10;
            byte[] SSN = new byte[Len_SSN];
            byte[] sHcaId = new byte[10];
            nErrCode = hpcGetHPCSSN(SSN, ref Len_SSN);
            if (nErrCode != 0)
            {
                msg("讀取錯誤。" + ErrCode.errMsg(nErrCode));
                return;
            }
            sHcaId = SSN;
            //病患基本資料
            byte[] sPatId = new byte[10];
            int buff = 72;
            byte[] pBuffer = new byte[buff];
            nErrCode = hisGetBasicData(pBuffer, ref buff);

            if (nErrCode != 0)
            {
                msg("讀取錯誤。" + ErrCode.errMsg(nErrCode));
                return;
            }
            string CardNo = BIG5.GetString(pBuffer, 0, 12).Trim();
            string Name = BIG5.GetString(pBuffer, 12, 20).Trim();
            string PID = BIG5.GetString(pBuffer, 32, 10).Trim();
            string Birthday = BIG5.GetString(pBuffer, 42, 7).Trim();
            string Gender = BIG5.GetString(pBuffer, 49, 1).Trim();
            string DeliverDate = BIG5.GetString(pBuffer, 50, 7).Trim();
            string VoidFlag = BIG5.GetString(pBuffer, 57, 1).Trim();
            string EmergencyPhoneNumber = BIG5.GetString(pBuffer, 58, 14).Trim();

            label1.Text = "卡片號碼：" + CardNo;
            label5.Text = "姓名：" + Name;
            label6.Text = "身分證號：" + PID;
            label7.Text = "出生日期：" + Birthday;
            label8.Text = "性別：" + Gender;
            label9.Text = "發卡日期：" + DeliverDate;
            label10.Text = "卡片註銷註記：" + VoidFlag;
            label11.Text = "緊急聯絡電話：" + EmergencyPhoneNumber;
            sPatId = System.Text.Encoding.Default.GetBytes(BIG5.GetString(pBuffer, 32, 10));//病患身分證

            byte[] sPatBirth = new byte[7];
            sPatBirth = System.Text.Encoding.Default.GetBytes(BIG5.GetString(pBuffer, 42, 7));//病患生日
            closeCom();
            nErrCode = PEA_SamExeNhiQuery(sHostName, nPort, sBusCode, nCom, sHcaId, sPatId, sPatBirth);

            //private static extern int PEA_GetMsg(byte[] sBuf,int nSize);
            int nSize = 256;
            byte[] sBuf = new byte[nSize];
            PEA_GetMsg(sBuf, ref nSize);
            string result = BIG5.GetString(sBuf);

            if (nErrCode == 0)
            {
                MessageBox.Show("用藥關懷名單有資料！！");
                p = System.Diagnostics.Process.Start("IExplore.exe", result);
                
            }
            else if (nErrCode == 1)
            {
                MessageBox.Show("查詢成功，無資料");
            }
            else if (nErrCode == -1)
            {
                MessageBox.Show("醫事機構端執行失敗");
            }
            else if (nErrCode == -2)
            {
                MessageBox.Show("健保局服務主機端錯誤");
            }

            rtOutput.Text += DateTime.Now.ToLongTimeString() + "：" + result;
            rtOutput.Text += Environment.NewLine;
            downTextbox();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openCom();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            closeCom();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            verifySAM();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            verifyHPC();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            getBasicData();
            hisReadPrescription();
            hisGetTreatmentNoNeedHPC();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            getBasicData();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            openCom();
            //讀卡機認證
            verifySAM();

            //醫事人員卡認證   
            verifyHPC();
            closeCom();
        }

        private void button9_Click(object sender, EventArgs e)
        {
             checkSAMStatus();
             checkHPCStatus();
             openCom();
             msg("安全模組狀態：" + hisGetCardStatus(1));
             msg("健保ic卡狀態：" + hisGetCardStatus(2));
             msg("醫事人員卡狀態：" + hisGetCardStatus(3));
             closeCom();

        }

        private void button8_Click(object sender, EventArgs e)
        {
            hisGetTreatmentNoNeedHPC();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            hisReadPrescription();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            //檢查健保卡狀態
            bool status = checkCardStatus();
            //如果非正常狀態，跳出訊息警告
            if (!status)
            {
                return;
            }
            //執行特定對象查詢動作
            PEA_SamExeNhiQuery();
        }

        private void rtOutput_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            p = System.Diagnostics.Process.Start("IExplore.exe", e.LinkText);
        }
    }
}
