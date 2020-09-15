using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Timers;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Configuration;

namespace COMPLETE_FLAT_UI
{
    public partial class FormMenuPrincipal : Form
    {
        //Constructor
        public FormMenuPrincipal()
        {
            InitializeComponent();
            //Những dòng này loại bỏ sự nhấp nháy của biểu mẫu hoặc điều khiển trong giao diện đồ họa (Nhưng không phải 100%)
            this.SetStyle(ControlStyles.ResizeRedraw, true);
            this.DoubleBuffered = true;
            LoadParameterFromConfigFile();
        }
        //PHƯƠNG PHÁP ĐẶT LẠI / THAY ĐỔI KÍCH THƯỚC ĐỂ ĐỊNH DẠNG THỜI GIAN CHẠY ----------------------------------------------------------
        private int tolerance = 15;
        private const int WM_NCHITTEST = 132;
        private const int HTBOTTOMRIGHT = 17;
        private Rectangle sizeGripRectangle;

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_NCHITTEST:
                    base.WndProc(ref m);
                    var hitPoint = this.PointToClient(new Point(m.LParam.ToInt32() & 0xffff, m.LParam.ToInt32() >> 16));
                    if (sizeGripRectangle.Contains(hitPoint))
                        m.Result = new IntPtr(HTBOTTOMRIGHT);
                    break;
                default:
                    base.WndProc(ref m);
                    break;
            }
        }
        //----------------DIBUJAR RECTANGULO / EXCLUIR ESQUINA PANEL 
        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            var region = new Region(new Rectangle(0, 0, this.ClientRectangle.Width, this.ClientRectangle.Height));

            sizeGripRectangle = new Rectangle(this.ClientRectangle.Width - tolerance, this.ClientRectangle.Height - tolerance, tolerance, tolerance);

            region.Exclude(sizeGripRectangle);
            this.panelContenedorPrincipal.Region = region;
            this.Invalidate();
        }
        //----------------COLOR Y GRIP DE RECTANGULO INFERIOR
        protected override void OnPaint(PaintEventArgs e)
        {

            SolidBrush blueBrush = new SolidBrush(Color.FromArgb(55, 61, 69));
            e.Graphics.FillRectangle(blueBrush, sizeGripRectangle);

            base.OnPaint(e);
            ControlPaint.DrawSizeGrip(e.Graphics, Color.Transparent, sizeGripRectangle);
        }
       
        //METODO PARA ARRASTRAR EL FORMULARIO---------------------------------------------------------------------
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void PanelBarraTitulo_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
        //METODOS PARA CERRAR,MAXIMIZAR, MINIMIZAR FORMULARIO------------------------------------------------------
        int lx, ly;
        int sw, sh;
        private void btnMaximizar_Click(object sender, EventArgs e)
        {
            lx = this.Location.X;
            ly = this.Location.Y;
            sw = this.Size.Width;
            sh = this.Size.Height;
            this.Size = Screen.PrimaryScreen.WorkingArea.Size;
            this.Location = Screen.PrimaryScreen.WorkingArea.Location;
            btnMaximizar.Visible = false;
            btnNormal.Visible = true;

        }

        private void btnNormal_Click(object sender, EventArgs e)
        {
            this.Size = new Size(sw, sh);
            this.Location = new Point(lx, ly);
            btnNormal.Visible = false;
            btnMaximizar.Visible = true;
        }

        private void btnMinimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc chắn đóng chương trình không?", "Cảnh Báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc chắn đóng chương trình không?", "Cảnh Báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        //METODOS PARA ANIMACION DE MENU SLIDING--
        private void btnMenu_Click(object sender, EventArgs e)
        {
            //-------CON EFECTO SLIDING
           /* if (panelMenu.Width == 230)
            {
                this.tmContraerMenu.Start();
            }
            else if (panelMenu.Width == 55)
            {
                this.tmExpandirMenu.Start();
            } */

            //-------SIN EFECTO 
            //if (panelMenu.Width == 55)
            //{
            //    panelMenu.Width = 230;
            //}
            //else

            //    panelMenu.Width = 55;
        }

        private void tmExpandirMenu_Tick(object sender, EventArgs e)
        {

            // if (panelMenu.Width >= 230)
            //     this.tmExpandirMenu.Stop();
            //else
            //    panelMenu.Width = panelMenu.Width + 5;
            ThreadOne();
            button2_Click(sender, e);
        }

        private void tmContraerMenu_Tick(object sender, EventArgs e)
        {
           /* if (panelMenu.Width <= 55)
                this.tmContraerMenu.Stop();
            else
                panelMenu.Width = panelMenu.Width - 5;
           */
        }

        //METODO PARA ABRIR FORM DENTRO DE PANEL-----------------------------------------------------
        private void AbrirFormEnPanel(object formHijo)
        {
            if (this.panelContenedorForm.Controls.Count > 0)
                this.panelContenedorForm.Controls.RemoveAt(0);
            Form fh = formHijo as Form;
            fh.TopLevel = false;
            fh.FormBorderStyle = FormBorderStyle.None;
            fh.Dock = DockStyle.Fill;            
            this.panelContenedorForm.Controls.Add(fh);
            this.panelContenedorForm.Tag = fh;
            fh.Show();
        }
        //METODO PARA MOSTRAR FORMULARIO DE LOGO Al INICIAR ----------------------------------------------------------
        private void MostrarFormLogo()
        {
            AbrirFormEnPanel(new FormLogo());
        }

        private void FormMenuPrincipal_Load(object sender, EventArgs e)
        {
            MostrarFormLogo();
            button2_Click(sender, e);
        }
        //METODO PARA MOSTRAR FORMULARIO DE LOGO Al CERRAR OTROS FORM ----------------------------------------------------------
        private void MostrarFormLogoAlCerrarForms(object sender, FormClosedEventArgs e)
        {
            MostrarFormLogo();
        }
        //METODOS PARA ABRIR OTROS FORMULARIOS Y MOSTRAR FORM DE LOGO Al CERRAR ----------------------------------------------------------
        private void btnListaClientes_Click(object sender, EventArgs e)
        {
            /*
            FormListaClientes fm = new FormListaClientes();
            fm.FormClosed += new FormClosedEventHandler(MostrarFormLogoAlCerrarForms);
            AbrirFormEnPanel(fm);*/
        }

        private void btnMembresia_Click(object sender, EventArgs e)
        {
            FormMembresia frm = new FormMembresia();
            frm.FormClosed += new FormClosedEventHandler(MostrarFormLogoAlCerrarForms);
            AbrirFormEnPanel(frm);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Formlistemail fm = new Formlistemail();
            fm.FormClosed += new FormClosedEventHandler(MostrarFormLogoAlCerrarForms);
            AbrirFormEnPanel(fm);
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void panelMenu_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelContenedorForm_Paint(object sender, PaintEventArgs e)
        {

        }

        //METODO PARA HORA Y FECHA ACTUAL ----------------------------------------------------------
        private void tmFechaHora_Tick(object sender, EventArgs e)
        {
            lbFecha.Text = DateTime.Now.ToLongDateString();
            lblHora.Text = DateTime.Now.ToString("HH:mm:ssss");
        }
        

        private void button5_Click(object sender, EventArgs e)
        {
            AbrirFormEnPanel(new Form1());
        }

        private static void ThreadOne()
        {
            Thread.Sleep(1000);
        }

        private void LoadParameterFromConfigFile()
        {
            try
            {
                
                //Load ORACLE connection string                    
                Common.ORACLE_SERVER =          GetConfigValue("ORACLE_SERVER");
                Common.ORACLE_PORT =            GetConfigValue("ORACLE_PORT");
                Common.ORACLE_SERVICE_NAME =    GetConfigValue("ORACLE_SERVICE_NAME");
                Common.ORACLE_USER =            GetConfigValue("ORACLE_USER");
                Common.ORACLE_PASSWORD =        GetConfigValue("ORACLE_PASSWORD");
                //load infor email 
                Common.SMTP_SERVER =            GetConfigValue("SMTP_SERVER");
                Common.SMTP_PORT =              GetConfigValue("SMTP_PORT");
                Common.SMTP_USERNAME =          GetConfigValue("SMTP_USERNAME");
                Common.SMTP_PASSWORD =          GetConfigValue("SMTP_PASSWORD");
                Common.MAIL_FROM =              GetConfigValue("MAIL_FROM");
                Common.MAIL_FROM_NAME =         GetConfigValue("MAIL_FROM_NAME");
               
            }
            catch (Exception)
            {
                MessageBox.Show("Error when loading config information. System will be down automatically. ", "Thông báo");
                Application.Exit();
            }

        }
        private string GetConfigValue(string configCode)
        {
            string strTemp = string.Empty;
            try
            {
 
                strTemp = ConfigurationManager.AppSettings[configCode].ToString();

                return strTemp;
            }
            catch (Exception)
            {
                MessageBox.Show("Error when loading config information. System will be down automatically. ", "Thông báo");
                return string.Empty;
            }

        }
    }
}
