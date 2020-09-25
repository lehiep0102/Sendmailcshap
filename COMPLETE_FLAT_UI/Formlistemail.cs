using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Common;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using Oracle.ManagedDataAccess.Client;
using System.Net;
using System.Globalization;

namespace MAS_EMAIL
{
    public partial class Formlistemail : Form
    {
        public string AppLocation = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
        public static bool checknull(string s)
        {
            return (s == null || string.IsNullOrEmpty(s) || s == " ") ? true : false;
        }
        public Formlistemail()
        {
            InitializeComponent();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormListaClientes_Load(object sender, EventArgs e)
        {
            InsertarFilas();
        }
        

        private void BtnCerrar_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnEditar_Click(object sender, EventArgs e)
        {
            Formdetails frm = new Formdetails();
            if (dataGridView1.SelectedRows.Count > 0)
            {               
                frm.txtid.Text= dataGridView1.CurrentRow.Cells[0].Value.ToString();
                frm.txtnombre.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                frm.txtapellido.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                frm.txtdireccion.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                frm.txttelefono.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                frm.ShowDialog();
            }
            else
                MessageBox.Show("vui lòng chọn một dòng dữ liệu");
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            int qr;
            int fille;
            int count = ckcount();
           // MessageBox.Show(count.ToString(), "path");
            if (count > 0)
            {
                fille = qrfill();
                qr = queryst();
            }
            else 
            {
                qr = queryst();
            }
           
        }

        private void InsertarFilas()
        {
           int qr = queryst();
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Formdetails frm = new Formdetails();
            if (dataGridView1.SelectedRows.Count > 0)
            {
                frm.txtid.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                frm.txtnombre.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                frm.txtapellido.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                frm.txtdireccion.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                frm.txttelefono.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                frm.ShowDialog();

            }
            else
                MessageBox.Show("vui lòng chọn một dòng dữ liệu");
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private int qrfill()
        {
            int count = 0;
            string acnt_no = "";
            string cust_nm = "";
            string birth_dt = " ";
            string ctry = "";
            string sex = "";
            string idno = "";
            string idno_iss_dt = " ";
            string idno_iss_orga = " ";
            string home_addr = " ";
            string office_addr = " ";
            string mobile = " ";
            string email = " ";
            string fax = " ";
            string rp = " ";
            string position = " ";
            string poa_no = " ";
            string dt_rp = " ";
            string stick = " ";
            string mtick = " ";
            string drtick = " ";
            string wtd = " ";
            string cct = " ";
            string acashad = " ";
            string sms_tk = "";
            string sotp = " ";
            string hotp = " ";
            string sr_otp = " ";
            string mtcar = " ";
            string sr_mt = " ";
            string ds = " ";
            string sr_ds = " ";
            string iss_by_ds = " ";
            string bank1 = " ";
            string bank_acc1 = " ";
            string bank_name1 = " ";
            string bank2 = " ";
            string bank_acc2 = " ";
            string bank_name2 = " ";
            string bank3 = " ";
            string bank_acc3 = " ";
            string bank_name3 = " ";
            string sc1 = " ";
            string acc_sc1 = " ";
            string note1 = " ";
            string sc2 = " ";
            string acc_sc2 = " ";
            string note2 = " ";
            string sc3 = " ";
            string acc_sc3 = " ";
            string note3 = " ";
            string sc4 = " ";
            string acc_sc4 = " ";
            string note4 = " ";

            OracleConnection conn = DBUtils.GetDBConnection();
            conn.Open();

            string sql = @"select   acnt_no,
                                    cust_nm,
                                    birth_dt,
                                    nvl(ctry,' '),
                                    nvl(sex,' '),
                                    nvl(idno,' '),
                                    nvl(idno_iss_dt,' '),
                                    nvl(idno_iss_orga,' '),
                                    nvl(home_addr,' '),
                                    nvl(office_addr,' '),
                                    nvl(mobile,' '),
                                    nvl(email,' '),
                                    nvl(fax,' '),
                                    nvl(rp,' '),
                                    nvl(position,' '),
                                    nvl(poa_no,' '),
                                    nvl(dt_rp,' '),
                                    nvl(stick,' '),
                                    nvl(mtick,' '),
                                    nvl(drtick,' '),
                                    nvl(wtd,' '),
                                    nvl(cct,' '),
                                    nvl(acashad,' '),
                                    nvl(sms_tk,' '),    
                                    nvl(sotp,' '),
                                    nvl(hotp,' '),
                                    nvl(sr_otp,' '),
                                    nvl(mtcar,' '),
                                    nvl(sr_mt,' '),
                                    nvl(ds,' '),
                                    nvl(sr_ds,' '),
                                    nvl(iss_by_ds,' '),
                                    nvl(bank1,' '),
                                    nvl(bank_acc1,' '),
                                    nvl(bank_name1,' '),
                                    nvl(bank2,' '),
                                    nvl(bank_acc2,' '),
                                    nvl(bank_name2,' '),
                                    nvl(bank3,' '),
                                    nvl(bank_acc3,' '),
                                    nvl(bank_name3,' '),
                                    nvl(sc1,' '),
                                    nvl(acc_sc1,' '),
                                    nvl(note1,' '),
                                    nvl(sc2,' '),
                                    nvl(acc_sc2,' '),
                                    nvl(note2,' '),
                                    nvl(sc3,' '),
                                    nvl(acc_sc3,' '),
                                    nvl(note3,' '),
                                    nvl(sc4,' '),
                                    nvl(acc_sc4,' '),
                                    nvl(note4,' ')
                                    from ACNT_CT where  FILL_YN = 'N'";

            try
            {
                OracleCommand cmd = new OracleCommand(sql, conn);// Tạo một đối tượng Command.

                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            acnt_no = reader.GetString(0);
                            cust_nm = reader.GetString(1);
                            birth_dt = reader.GetString(2);
                            ctry = reader.GetString(3);
                            sex = reader.GetString(4);
                            idno = reader.GetString(5);
                            idno_iss_dt = reader.GetString(6);
                            idno_iss_orga = reader.GetString(7);
                            home_addr = reader.GetString(8);
                            office_addr = reader.GetString(9);
                            mobile = reader.GetString(10);
                            email = reader.GetString(11);
                            fax = reader.GetString(12);
                            rp = reader.GetString(13);
                            position = reader.GetString(14);
                            poa_no = reader.GetString(15);
                            dt_rp = reader.GetString(16);
                            stick = reader.GetString(17);
                            mtick = reader.GetString(18);
                            drtick = reader.GetString(19);
                            wtd = reader.GetString(20);
                            cct = reader.GetString(21);
                            acashad = reader.GetString(22);
                            sms_tk = reader.GetString(23);
                            sotp = reader.GetString(24);
                            hotp = reader.GetString(25);
                            sr_otp = reader.GetString(26);
                            mtcar = reader.GetString(27);
                            sr_mt = reader.GetString(28);
                            ds = reader.GetString(29);
                            sr_ds = reader.GetString(30);
                            iss_by_ds = reader.GetString(31);
                            bank1 = reader.GetString(32);
                            bank_acc1 = reader.GetString(33);
                            bank_name1 = reader.GetString(34);
                            bank2 = reader.GetString(35);
                            bank_acc2 = reader.GetString(36);
                            bank_name2 = reader.GetString(37);
                            bank3 = reader.GetString(38);
                            bank_acc3 = reader.GetString(39);
                            bank_name3 = reader.GetString(40);
                            sc1 = reader.GetString(41);
                            acc_sc1 = reader.GetString(42);
                            note1 = reader.GetString(43);
                            sc2 = reader.GetString(44);
                            acc_sc2 = reader.GetString(45);
                            note2 = reader.GetString(46);
                            sc3 = reader.GetString(47);
                            acc_sc3 = reader.GetString(48);
                            note3 = reader.GetString(49);
                            sc4 = reader.GetString(50);
                            acc_sc4 = reader.GetString(51);
                            note4 = reader.GetString(52);
                            string datafill = filldatadocx(acnt_no, cust_nm, birth_dt, ctry, sex, idno, idno_iss_dt, idno_iss_orga, home_addr, office_addr, mobile, email, fax, rp, position, poa_no, dt_rp, stick, mtick, drtick, wtd, cct, acashad, sms_tk, sotp, hotp, sr_otp, mtcar, sr_mt, ds, sr_ds, iss_by_ds, bank1, bank_acc1, bank_name1, bank2, bank_acc2, bank_name2, bank3, bank_acc3, bank_name3, sc1, acc_sc1, note1, sc2, acc_sc2, note2, sc3, acc_sc3, note3, sc4, acc_sc4, note4);
                            count = count + 1;
                        }
                    }
                }
            }
            catch (Exception)
            {
                return 1;
            }
            finally
            {

                conn.Close();
                conn.Dispose();
                conn = null;
            }

            if (count == 0)
            {
                return 1;
            }
            return 0;
        }
        private string filldatadocx(string acnt_no, string cust_nm, string birth_dt, string ctry, string sex, string idno, string idno_iss_dt, string idno_iss_orga, string home_addr, string office_addr, string mobile, string email, string fax, string rp, string position, string poa_no, string dt_rp, string stick, string mtick, string drtick, string wtd, string cct, string acashad, string sms_tk, string sotp, string hotp, string sr_otp, string mtcar, string sr_mt, string ds, string sr_ds, string iss_by_ds, string bank1, string bank_acc1, string bank_name1, string bank2, string bank_acc2, string bank_name2, string bank3, string bank_acc3, string bank_name3, string sc1, string acc_sc1, string note1, string sc2, string acc_sc2, string note2, string sc3, string acc_sc3, string note3, string sc4, string acc_sc4, string note4)
        {
            char tick = '\u2611';
            char nontick = '\u2B1C';
            string nullid = "…………………….";
            AppLocation = AppLocation.Replace("file:\\", "");
            Object oMissing = System.Reflection.Missing.Value;
            Object oTemplatePath = AppLocation + "\\Request_Contract_opening_11March2020.dotx";//global::MAS_EMAIL.Properties.Resources.Request_Contract_opening_11March2020;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
          //  MessageBox.Show(oTemplatePath.ToString());
            Document wordDoc = new Document();
            wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            foreach (Field myMergeField in wordDoc.Fields)
            {

                Range rngFieldCode = myMergeField.Code;

                String fieldText = rngFieldCode.Text;

                if (fieldText.StartsWith(" MERGEFIELD"))
                {
                    // THE TEXT COMES IN THE FORMAT OF
                    // MERGEFIELD  MyFieldName  \\* MERGEFORMAT
                    // THIS HAS TO BE EDITED TO GET ONLY THE FIELDNAME "MyFieldName"

                    Int32 endMerge = fieldText.IndexOf("\\");

                    Int32 fieldNameLength = fieldText.Length - endMerge;

                    String fieldName = fieldText.Substring(11, endMerge - 11);

                    // GIVES THE FIELDNAMES AS THE USER HAD ENTERED IN .dot FILE

                    fieldName = fieldName.Trim();
                    //MessageBox.Show(fieldName);
                    if (fieldName == "1")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(acnt_no.Substring(0, 1));
                    }

                    if (fieldName == "2")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(acnt_no.Substring(1, 1));
                    }

                    if (fieldName == "3")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(acnt_no.Substring(2, 1));
                    }

                    if (fieldName == "4")
                    {
                        myMergeField.Select();
                        // MessageBox.Show(acnt_no.Substring(4, 1));
                        wordApp.Selection.TypeText(acnt_no.Substring(3, 1));
                    }

                    if (fieldName == "5")
                    {
                        myMergeField.Select();
                        //  MessageBox.Show(acnt_no.Substring(5, 1));
                        wordApp.Selection.TypeText(acnt_no.Substring(4, 1));
                    }

                    if (fieldName == "6")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(acnt_no.Substring(5, 1));
                    }
                    if (fieldName == "7")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(acnt_no.Substring(6, 1));
                    }

                    if (fieldName == "8")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(acnt_no.Substring(7, 1));
                    }

                    if (fieldName == "9")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(acnt_no.Substring(8, 1));
                    }

                    if (fieldName == "10")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(acnt_no.Substring(9, 1));
                    }

                    if (fieldName == "name")
                    {

                        myMergeField.Select();
                        //MessageBox.Show(cust_nm);
                        wordApp.Selection.TypeText(cust_nm);

                    }

                    if (fieldName == "birthdt")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(birth_dt.Substring(6, 2) + "/" + birth_dt.Substring(4, 2) + "/" + birth_dt.Substring(0, 4));
                        //MessageBox.Show(birth_dt.Substring(6, 2) + "/" + birth_dt.Substring(4, 2) + "/" + birth_dt.Substring(0, 4));
                    }

                    if (fieldName == "nt")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(ctry);

                    }

                    if (fieldName == "sex")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(sex);

                    }

                    if (fieldName == "idnb")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(idno);

                    }

                    if (fieldName == "dtis")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(idno_iss_dt.Substring(6, 2) + "/" + idno_iss_dt.Substring(4, 2) + "/" + idno_iss_dt.Substring(0, 4));

                    }
                    if (fieldName == "poiss")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(idno_iss_orga);

                    }

                    if (fieldName == "cradd")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(home_addr);

                    }


                    if (fieldName == "wplace")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(office_addr);

                    }

                    if (fieldName == "phone")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(mobile);

                    }
                    if (fieldName == "email")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(email);

                    }
                    if (fieldName == "fax")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(fax);

                    }
                    if (fieldName == "rp")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(rp);

                    }
                    if (fieldName == "position")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(position);

                    }
                    if (fieldName == "poa_no")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(poa_no);

                    }
                    if (fieldName == "dt_rp")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(dt_rp);

                    }
                    if (fieldName == "ckcst")
                    {
                        if (stick == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }


                    }
                    if (fieldName == "mrgin")
                    {
                        if (mtick == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }


                    }
                    if (fieldName == "ckps")
                    {
                        if (drtick == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }
                    }
                    if (fieldName == "td")
                    {
                        if (wtd == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }
                    }
                    if (fieldName == "cctd")
                    {
                        if (cct == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }
                    }
                    if (fieldName == "uttbck")
                    {
                        if (acashad == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }
                    }

                    if (fieldName == "smspp")
                    {
                        if (sms_tk == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }
                    }
                    if (fieldName == "sotp")
                    {
                        if (sotp == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }
                    }
                    if (fieldName == "hotp")
                    {
                        if (hotp == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }
                    }
                    if (fieldName == "serihotp")
                    {
                        if (hotp == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(sr_otp);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(nullid);
                        }
                    }

                    if (fieldName == "mtcar")
                    {
                        if (mtcar == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }
                    }
                    if (fieldName == "mtsr")
                    {
                        if (mtcar == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(sr_mt + "…………...");
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(nullid);
                        }
                    }

                    if (fieldName == "ckso")
                    {
                        if (ds == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(tick));
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(char.ToString(nontick));
                        }
                    }

                    if (fieldName == "srcks")
                    {
                        if (ds == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(sr_ds);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(nullid);
                        }
                    }

                    if (fieldName == "issby")
                    {
                        if (ds == "Y")
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(iss_by_ds);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(nullid);
                        }
                    }

                    if (fieldName == "stt1")
                    {
                        if (checknull(bank1) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText("1");
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "tenctk1")
                    {
                        if (checknull(bank1) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(bank_name1);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "stk1")
                    {
                        if (checknull(bank1) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(bank_acc1);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }

                    if (fieldName == "bank1")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(bank1);
                    }

                    if (fieldName == "stt2")
                    {
                        if (checknull(bank2) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText("2");
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "tenctk2")
                    {
                        if (checknull(bank2) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(bank_name2);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "stk2")
                    {
                        if (checknull(bank2) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(bank_acc2);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }

                    if (fieldName == "bank2")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(bank2);
                    }

                    if (fieldName == "stt3")
                    {
                        if (checknull(bank3) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText("3");
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "tentk3")
                    {
                        if (checknull(bank3) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(bank_name3);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "stk3")
                    {
                        if (checknull(bank3) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(bank_acc3);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }

                    if (fieldName == "bank3")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(bank3);
                    }

                    if (fieldName == "ckstt1")
                    {
                        if (checknull(sc1) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText("1");
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "stkck1")
                    {
                        if (checknull(sc1) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(acc_sc1);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "noteck1")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(note1);

                    }

                    if (fieldName == "ctck1")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(sc1);
                    }

                    if (fieldName == "ckstt2")
                    {
                        if (checknull(sc2) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText("2");
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "stkck2")
                    {
                        if (checknull(sc2) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(acc_sc2);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "noteck2")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(note2);

                    }

                    if (fieldName == "ctck2")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(sc2);
                    }

                    if (fieldName == "ckstt3")
                    {
                        if (checknull(sc3) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText("3");
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "stkck3")
                    {
                        if (checknull(sc3) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(acc_sc3);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "noteck3")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(note3);

                    }

                    if (fieldName == "ctck3")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(sc3);
                    }

                    if (fieldName == "ckstt4")
                    {
                        if (checknull(sc4) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText("4");
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "stkck4")
                    {
                        if (checknull(sc4) == false)
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(acc_sc4);
                        }
                        else
                        {
                            myMergeField.Select();
                            wordApp.Selection.TypeText(" ");
                        }
                    }
                    if (fieldName == "noteck4")
                    {

                        myMergeField.Select();
                        wordApp.Selection.TypeText(note4);

                    }

                    if (fieldName == "ctck4")
                    {
                        myMergeField.Select();
                        wordApp.Selection.TypeText(sc4);
                    }

                }
            }

            wordDoc.SaveAs(AppLocation + "\\"+acnt_no + ".docx");
            wordDoc.SaveAs2(AppLocation + "\\" + acnt_no + ".pdf", WdSaveFormat.wdFormatPDF, oMissing, oMissing, oMissing,oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            //wordApp.Documents.Open("myFile.docx");
            wordDoc.Close();
            wordApp.Application.Quit();

            int update = updatefill(acnt_no);
            //string returnkq = "fill OK";
            return "fill OK"; 
        }
        private int updatefill(string acnt_no) // cập nhật trạng thái tài khoản này đã fill xong
        {
            OracleConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            String countinser = "";
            try
            {
                string sql = "update ACNT_CT set FILL_YN = 'Y' where ACNT_NO = trim('" + acnt_no + "') and  acnt_no = trim('" + acnt_no + "')";
                OracleCommand cmd = new OracleCommand(sql, conn);// Tạo một đối tượng Command.
                int rowCount = cmd.ExecuteNonQuery();
                countinser = rowCount.ToString();
            }
            catch (Exception)
            {
                return 1;
            }
            finally
            {

                conn.Close();
                conn.Dispose();
                conn = null;
                insertpdf(acnt_no);
            }
            return 0;
        }
        private string insertpdf(string acnt_nos)// update nội dung pdf , update data preview HĐ, update ID vảo table ANCT_ON
        {
            //AppLocation = AppLocation.Replace("file:\\", "");
            ///byte[] array;
            string fileName = "\\"+acnt_nos + ".pdf";
            string filepath = AppLocation + fileName;
            string fileType = "contract_"+acnt_nos + ".pdf";
            //OracleParameter param = new OracleParameter();
            //OracleConnection conn = DBUtils.GetDBConnection();
            //conn.Open();
            // string str = "";
            // string countinser = "";
            string sid = Guid.NewGuid().ToString().ToUpper();
            string sdatalink = "";

            using (OracleConnection oc = DBUtils.GetDBConnection())
            {
                oc.Open();
                string sData = Convert.ToBase64String(File.ReadAllBytes(filepath));
                OracleCommand cmd = new OracleCommand("INSERT INTO ACNT_PDF values (:1, :2, :3, :4, SYSDATE,SYSDATE,SYSDATE,'VN','1')", oc);
                cmd.Parameters.Add(new OracleParameter("1", OracleDbType.NVarchar2, acnt_nos, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("2", OracleDbType.NVarchar2, fileType, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("3", OracleDbType.NVarchar2, sData.Length, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("4", OracleDbType.Clob, sData, ParameterDirection.Input));
                cmd.ExecuteNonQuery();
                // insert data preview HĐ
                sdatalink = "<Document><Data><Static><Id>" + sid + "</ Id >< Name >" + fileType + "</ Name ></Static><Dynamic><Content>" + sData + "</Content></Dynamic></Data><Document>";
                OracleCommand cmd2 = new OracleCommand("INSERT INTO DATA_LINK values (:1, :2, :3)", oc);
                cmd2.Parameters.Add(new OracleParameter("1", OracleDbType.NVarchar2, sid, ParameterDirection.Input));
                cmd2.Parameters.Add(new OracleParameter("2", OracleDbType.Clob, sdatalink, ParameterDirection.Input));
                cmd2.Parameters.Add(new OracleParameter("3", OracleDbType.NVarchar2, fileType, ParameterDirection.Input));
                cmd2.ExecuteNonQuery();

                //update ID thông tin preview HĐ cho bảng master
                OracleCommand cmd3 = new OracleCommand("update ACNT_ON set IDCT = :1 where acnt_no = substr(:2,4,7)", oc);
                cmd3.Parameters.Add(new OracleParameter("1", OracleDbType.NVarchar2, sid, ParameterDirection.Input));
                cmd3.Parameters.Add(new OracleParameter("2", OracleDbType.NVarchar2, acnt_nos, ParameterDirection.Input));
                cmd3.ExecuteNonQuery();

                //MessageBox.Show("fill data done");
                if (File.Exists(Path.Combine(filepath)))
                {
                    File.Delete(Path.Combine(filepath));
                    File.Delete(Path.Combine(AppLocation +"\\"+ acnt_nos + ".docx"));
                }
                oc.Close();
            }
            using (OracleConnection real = MAS_EMAIL.DBUtilsreal.GetDBConnection())
            {
                real.Open();
                OracleCommand cmdreal = new OracleCommand("INSERT INTO DATA_LINK values (:1, :2, :3)", real);
                cmdreal.Parameters.Add(new OracleParameter("1", OracleDbType.NVarchar2, sid, ParameterDirection.Input));
                cmdreal.Parameters.Add(new OracleParameter("2", OracleDbType.Clob, sdatalink, ParameterDirection.Input));
                cmdreal.Parameters.Add(new OracleParameter("3", OracleDbType.NVarchar2, fileType, ParameterDirection.Input));
                cmdreal.ExecuteNonQuery();
                real.Close();
                real.Dispose();
            }
            string get = getpdf(acnt_nos);

            if (get == "0")
            {
                return "1";
            }
            // string refsa = "Done";
            return "0";
        }
        private string getpdf(string acnt_nos) //get pdf để send mail 
        {
           // AppLocation = AppLocation.Replace("file:\\", "");
            OracleConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            string bytes;
            string fileName;
            string pathfile;
            int createm;

            try
            {
                string sql = @"select acnt_no,data from ACNT_PDF t where acnt_no  = '" + acnt_nos + "'";

                OracleCommand cmd = new OracleCommand(sql, conn);
                // Liên hợp Command với Connection.
                using (OracleDataReader sdr = cmd.ExecuteReader())
                {
                    sdr.Read();
                    fileName = sdr["acnt_no"].ToString();
                    bytes = sdr["Data"].ToString();
                }
                
                pathfile = AppLocation +"\\data\\"+"contract_" + fileName + ".pdf";
                //MessageBox.Show(pathfile, "path");
                using (System.IO.FileStream stream = System.IO.File.Create(pathfile))
                {
                    System.Byte[] byteArray = System.Convert.FromBase64String(bytes);
                    stream.Write(byteArray, 0, byteArray.Length);
                    //System.Diagnostics.Process.Start("explorer.exe", string.Format("/select,\"{0}\"", pathfile));
                }
            }
            catch (Exception)
            {
                return "1";
            }
            finally
            {
                conn.Close();
                conn.Dispose();
                conn = null;
                createm = createemail(acnt_nos);
            }

           // string refsa = "Done";
           if (createm!=0)
            {
                return "1";
            }
            return "0";
        }
        private int createemail(string acnt_no)
        {
            OracleConnection conn = DBUtils.GetDBConnection();

            IPAddress iddress = ip.GetLocalIPAddress();
            string user = "SYS";
            int sem;

            try
            {
                OracleCommand cmd = new OracleCommand("PACC_EM", conn);
                OracleTransaction transaction;
                conn.Open();
                // Start a local transaction
                transaction = conn.BeginTransaction(IsolationLevel.ReadCommitted);
                // Assign transaction object for a pending local transaction
                cmd.Transaction = transaction;

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("i_acnt_no", OracleDbType.Varchar2).Value = acnt_no;
                cmd.Parameters.Add("is_user_id", OracleDbType.Varchar2).Value = user;
                cmd.Parameters.Add("is_ip", OracleDbType.Varchar2).Value = iddress;
                cmd.ExecuteNonQuery();
                transaction.Commit();
                transaction.Dispose();
                cmd.Dispose();
            }
            catch (Exception)
            {
                // MessageBox.Show("lỗi" + ea.Message, "Cảnh báo");
                return 1;

            }
            finally
            {

                conn.Close();
                conn.Dispose();
                conn = null;
                sem = sendaccon(acnt_no);
            }
            if (sem != 0)
            {
                return 1;
            }
            return 0;
        }

        private int sendaccon(string acnt_no)
        {
            AppLocation = AppLocation.Replace("file:\\", "");
            //string scc = "";
            string sescc = "";
            string acnt_noem;
            string cust_nm;
            string sub;
            string email_bd;
            string[] file = new string[1];
            string[] mailsend = new string[1];
            string[] mailcc = new string[0];
            string[] mailbcc = new string[0];
            string pathfile = AppLocation+"\\data\\" + "contract_" + acnt_no + ".pdf";

            OracleConnection conn = DBUtils.GetDBConnection();
            conn.Open();

            string sql = "select t.acnt_no, t.cust_nm, t.email,t.sub,t.email_bd from ACNT_EMAIL t where t.send_yn <> 'Y' and t.acnt_no like'" + acnt_no + "%'";

            OracleCommand cmd = new OracleCommand(sql, conn);// Tạo một đối tượng Command.
            using (DbDataReader reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        acnt_noem = reader.GetString(0);
                        cust_nm = reader.GetString(1);
                        mailsend[0] = reader.GetString(2);
                        sub = reader.GetString(3);
                        email_bd = reader.GetString(4);
                        file[0] = pathfile;
                        /* Send email */
                        emailsend emailsn = new emailsend();
                        sescc = emailsn.SendEmail(mailsend, mailcc, mailbcc, sub, email_bd, file, acnt_noem, true);
 
                        if (sescc == "0")
                        {
                            int up = updatesenmail(acnt_noem);
                        }
                        else
                        {
                            int err = uperorsenmail(acnt_noem);
                        }
                    }
                }
            }
            conn.Close();
            conn.Dispose();
            return 0;
        }

        private int updatesenmail(string acnt_no)
        {
            OracleConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            String countinser = "";
            try
            {
                string sql = "update ACNT_EMAIL set SEND_YN = 'Y',DT_S = sysdate where acnt_no ='" + acnt_no + "'";
                OracleCommand cmd = new OracleCommand(sql, conn);// Tạo một đối tượng Command.
                int rowCount = cmd.ExecuteNonQuery();
                countinser = rowCount.ToString();
            }
            catch (Exception)
            {
                return 0;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
                conn = null;
                int ar = queryst();
            }
            return 1;
        }
        private int uperorsenmail(string acnt_no)
        {
            OracleConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            String countinser = "";
            try
            {
                string sql = "update ACNT_EMAIL set SEND_YN = 'E',DT_S = sysdate where acnt_no ='" + acnt_no + "'";
                OracleCommand cmd = new OracleCommand(sql, conn);// Tạo một đối tượng Command.
                int rowCount = cmd.ExecuteNonQuery();
                countinser = rowCount.ToString();
            }
            catch (Exception)
            {
                return 1 ;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
                conn = null;
                int ar = queryst();
            }
            return 0;
        }

        private int queryst()
        {
            OracleConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            try
            {
                string sql = "select t.send_yn,t.acnt_no,t.cust_nm,t.email,t.dt_s from ACNT_EMAIL t order by t.work_dtm desc";

                OracleCommand cmd = new OracleCommand();
                // Liên hợp Command với Connection.
                cmd.Connection = conn;
                cmd.CommandText = sql;

                System.Data.DataTable dataemail = new System.Data.DataTable();
                dataemail.Columns.Add("Trạng Thái", typeof(string));
                dataemail.Columns.Add("Số TK", typeof(string));
                dataemail.Columns.Add("Tên Khách hàng", typeof(string));
                dataemail.Columns.Add("Email", typeof(string));
                dataemail.Columns.Add("Ngày giờ gửi email", typeof(DateTime));

                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            string[] array_msg = new string[5];
                            array_msg[0] = reader.GetString(0);
                            array_msg[1] = reader.GetString(1);
                            array_msg[2] = reader.GetString(2);
                            array_msg[3] = reader.GetString(3);
                            array_msg[4] = reader.GetDateTime(4).ToString();
                            dataemail.Rows.Add(array_msg);
                        }

                        DataSet dsDataset = new DataSet();

                        dsDataset.Tables.Add(dataemail);
                        dataGridView1.DataSource = "";
                        dataGridView1.DataSource = dsDataset.Tables[0];
                    }
                }

                var culture = CultureInfo.CreateSpecificCulture("en-GB");
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Times New Roman", 11, System.Drawing.FontStyle.Bold);
                dataGridView1.DefaultCellStyle.Font = new System.Drawing.Font("Times New Roman", 11);
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[0].Width = 150;
                dataGridView1.Columns[1].Width = 180;
                dataGridView1.Columns[2].Width = 240;
                dataGridView1.Columns[3].Width = 180;
                dataGridView1.Columns[4].Width = 240;
                dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            }
            catch (Exception)
            {
                return 1;// MessageBox.Show("Tra cứu dữ liệu không thành công" + ea.Message, "Cảnh báo");
            }
            finally
            {
                conn.Close();
                conn.Dispose();
                conn = null;
            }

            string sent = "Đã gửi";
            string nsent = "Chưa gửi";
            string esent = "Gửi lỗi";

            for (int index = 0; index < dataGridView1.Rows.Count - 1; index++)
            {
                string status = dataGridView1.Rows[index].Cells[0].Value.ToString();
                if (status == "Y")
                {
                    dataGridView1.Rows[index].Cells[0].Value = sent;
                    dataGridView1.Rows[index].Cells[0].Style.BackColor = System.Drawing.Color.GreenYellow;
                    dataGridView1.Rows[index].Cells[0].Style.ForeColor = System.Drawing.Color.Black;
                    dataGridView1.Rows[index].Cells[0].Style.Font = new System.Drawing.Font("Times New Roman", 12);
                }
                else if (status == "E")
                {
                    dataGridView1.Rows[index].Cells[0].Value = esent;
                    dataGridView1.Rows[index].Cells[0].Style.BackColor = System.Drawing.Color.Red;
                    dataGridView1.Rows[index].Cells[0].Style.ForeColor = System.Drawing.Color.White;
                    dataGridView1.Rows[index].Cells[0].Style.Font = new System.Drawing.Font("Times New Roman", 12);
                }
                else
                {
                    dataGridView1.Rows[index].Cells[0].Value = nsent;
                    dataGridView1.Rows[index].Cells[0].Style.BackColor = System.Drawing.Color.Orange;
                    dataGridView1.Rows[index].Cells[0].Style.ForeColor = System.Drawing.Color.Black;
                    dataGridView1.Rows[index].Cells[0].Style.Font = new System.Drawing.Font("Times New Roman", 12);
                }
            }
            return 0;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            btnNuevo_Click(sender,e);
        }
        private int ckcount()
        {
            int rowCount=0;
            OracleConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            string sql = @"select count(*) count from ACNT_CT where  FILL_YN = 'N'";
            OracleCommand cmd = new OracleCommand(sql, conn);
            try
            {
                OracleDataReader sdr = cmd.ExecuteReader();
                sdr.Read();
                rowCount = Int32.Parse(sdr["count"].ToString());
            }
            catch
            {
                return 0;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
                conn = null;
            }

            if (rowCount > 0)
            {
                return rowCount;
            }
            else
                return 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string emailresen = emailrs.Text;
            int send = sendaccon(emailresen);
            /*AppLocation = AppLocation.Replace("file:\\", "");
            string[] file = new string[1];
            string[] mailsend = new string[1];
            string[] mailcc = new string[0];
            string[] mailbcc = new string[0];
            string sub = "test email SLL";
            string email_bd = "adasdasdasdsadasdsadsadsadsadasdsadasdsadsadsadasd";
            string sescc = "";
            string acnt_no = "077C112010";
            string pathfile = AppLocation + "contract_" + acnt_no + ".pdf";
            string acnt_noem = acnt_no;
            file[0] = pathfile;
            mailsend[0] = "hiepvanle92@gmail.com";

            SendEcEmail emailsn = new SendEcEmail();
            sescc = emailsn.SendEmail(mailsend, mailcc, mailbcc, sub, email_bd, file, acnt_noem, true);
            */
        }
    }
}
