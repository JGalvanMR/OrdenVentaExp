using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using MySql.Data.Types;
using System.Globalization;
using System.IO;
using System.Drawing.Printing;
using System.Diagnostics;
using System.Data.OleDb;
using System.Linq.Expressions;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using TextBox = System.Windows.Forms.TextBox;
using Font = System.Drawing.Font;
using Application = System.Windows.Forms.Application;


namespace OrdenVentaExp
{
    public partial class OrdVenExp : Form
    {
        //SqlConnection thisConnection2 = new SqlConnection("Data Source= gab.dyndns.org,2351;Initial Catalog=UERP_Irapuato;Connect Timeout=130;User ID=sa; MultipleActiveResultSets=True");               
        //SqlConnection thisConnection2 = new SqlConnection("Data Source=201.120.60.21,2351;Initial Catalog=UERP_Irapuato;Connect Timeout=130;User ID=sa; MultipleActiveResultSets=True");
        SqlConnection thisConnection2 = new SqlConnection(Utilerias.Class1.ConnectionStringUERP);
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlConnection thisConnecTomate = new SqlConnection(Utilerias.Class1.ConnectionAgui);
        //SqlConnection thisConnecTomate = new SqlConnection("Data Source= pronase1.ddns.net\\SQL2022,2359;Initial Catalog=GAB_Empaque;Connect Timeout=130;User ID=SipGab; Password=Empaque1$; MultipleActiveResultSets=True");
        //OleDbConnection MyConnection = new OleDbConnection(@"Provider= VFPOLEDB.1;Data Source=c:\mr_lucky\base_de_datos;Collating Sequence=general;");
        //OleDbConnection MyConnection = new OleDbConnection(@"Provider= VFPOLEDB.1;Data Source=\\gabira1\mr_lucky\base_de_datos;Collating Sequence=general;");
        SqlCommand cmnd2, cmnd3, cmnd4, cmnd6;
        SqlDataReader reader2, reader3, reader4, reader6;
        DataTable tmgrdped = new DataTable("tmgrdped");
        DataTable entradas = new DataTable("entradas");
        DataTable subcnte = new DataTable("subcnte");
        DataTable tarimas = new DataTable("tarimas");
        DataTable CatUser = new DataTable();
        DataTable Fraccion = new DataTable();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        SqlDataAdapter da = new SqlDataAdapter();
        TextBox txt1 = new TextBox();
        string ConUERP = "";
        string ExisteUERP = "";
        //OleDbCommand cmd1 = new OleDbCommand();

        public static string lin_clave, prod_clave, prod_nombre, cnte_sucursal, cnte_entregas, pdn_folio, ArtCod = " ", ArtNom, cveunidad = " ", item_number = "";
        public static string sig, id_cliente, id_vendedor, um_clave, ultimofolio, cve_subcli, moneda, pdn_folio_web, vend_clave = "", vend_nombre;
        public static string ryandigital = " ", ryantipo, pdn_pedimento, pdn_fechap, pdn_aduana, pdn_pre_cad, pdn_cantxcaja, pdn_tipo, pdn_estatus;
        public static string flete_gab = "N", monedaflete = "", tipoflete = "", iva_flete = "N", especial_venta = "";
        public static string TA = "", TB = "", TC = "", TD = "", TE = "", tme = "", tem = "", tempo = "", leyenda1 = "", destino = "", CveSubCli = "", TipoPed = "";
        public static decimal precio = 0, preciouni, resulta, prod_num_tarimas, sumataraprox, sumacomp, sumasplit, enteros, decimales;
        public static decimal prod_precio_mn, prod_precio_usd, var_peso_unitario, var_dec_peso_total, pdn_num_unidades, pdn_precio_mn;
        public static decimal pdn_precio_usd, moneyusd, moneymn, moneyeuro;
        public static bool encontro = false;
        public static int sige, cantidad, sumapedido, opcion, pedorigen = 0, CantMod = 0;
        ComboBox vmenunombre = new ComboBox();
        ComboBox vmenuclave = new ComboBox();
        public DataTable Faltantes = new DataTable();
        bool estadoimp = false;
        int cont_prod = 1, ultimo_renglon = 0;

        string Actualizar = "N";

        private bool requiereActualizar = false;

        //crea la tabla del DGV
        private void MakeDataTabletmgrdhis()
        {
            tmgrdped.Columns.Add("linea");
            tmgrdped.Columns.Add("producto");
            tmgrdped.Columns.Add("nombre");
            tmgrdped.Columns.Add("precio");
            tmgrdped.Columns.Add("cantidad");
            tmgrdped.Columns.Add("clave_tarima");
            tmgrdped.Columns.Add("nom_tarima");
            tmgrdped.Columns.Add("subcnte_clave");
            tmgrdped.Columns.Add("subcliente");
            tmgrdped.Columns.Add("po_number");
            tmgrdped.Columns.Add("total");
            tmgrdped.Columns.Add("observaciones");
            tmgrdped.Columns.Add("estatus");
            tmgrdped.Columns.Add("transito");
            tmgrdped.Columns.Add("especial");
            tmgrdped.Columns.Add("special");
            tmgrdped.Columns.Add("fecha_cad");
            tmgrdped.Columns.Add("pedimento");
            DataColumn[] primarykeycolumn = new DataColumn[1];
            primarykeycolumn[0] = tmgrdped.Columns["producto"];
            tmgrdped.PrimaryKey = primarykeycolumn;
            ds1.Tables.Add(tmgrdped);
        }

        //carga todos los productos del catalogo
        private void cat_productos()
        {
            #region combo box productos
            CBProducto.Items.Clear();
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "select tmvc.prod_clave, tcpr.prod_nombre from tb_mstr_venta_compartida tmvc, tb_cat_producto tcpr where tmvc.vcp_tipo_venta='EXP' and tcpr.prod_clave = tmvc.prod_clave AND (tcpr.PROD_TIPO='PTP' OR tcpr.PROD_TIPO='PTC') order by tcpr.prod_nombre";
            reader2 = cmnd2.ExecuteReader();
            while (reader2.Read())
            {
                CBProducto.Items.Add(reader2.GetSqlString(1).ToString().Trim());
            }
            reader2.Close();
            #endregion
        }

        //carga los proveedores
        private void transportista()
        {
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "select distinct prov_nombre from tb_cat_proveedor order by prov_nombre";
            reader2 = cmnd2.ExecuteReader();
            while (reader2.Read())
            {
                CBTransportista.Items.Add(reader2.GetSqlString(0).ToString());
            }
            reader2.Close();
        }

        public OrdVenExp()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
            LblFecArchivo.Text = "Ver. " + Convert.ToDateTime(File.GetLastWriteTime(@"C:\Sisgabweb\OrdenVentaExp.exe")).ToString();
        }

        private void PBbtnimpre_Click(object sender, EventArgs e)
        {
            Impre impresora = new Impre();
            impresora.ShowDialog();
        }

        //habilitar controles
        private void habilita(bool estado)
        {
            CBmoneda.Enabled = estado;
            CBCliente.Enabled = estado;
            txtcliente.Enabled = estado;
            txtobs.Enabled = estado;
            txtpedsigma.Enabled = estado;
            txttalonemb.Enabled = estado;
            txtimporte.Enabled = estado;
            CBTransportista.Enabled = estado;
            txtclavetrans.Enabled = estado;
            txtcajatrailer.Enabled = estado;
            txtsello.Enabled = estado;
            txtsubtotal.Enabled = estado;
            txtIVAfle.Enabled = estado;
            txtretension.Enabled = estado;
            cbIVA.Enabled = estado;
            CBmonfle.Enabled = estado;
            cbpagoGAB.Enabled = estado;
            CBtipoflete.Enabled = estado;
            txtcompleped.Enabled = estado;
            txtmn.Enabled = estado;
            txtfecha.Enabled = estado;
            cbdestino.Enabled = estado;
            DtHrCita.Enabled = estado;
            TxtTractor.Enabled = estado;
            TxtCjTracto.Enabled = estado;
            TxtCjTra.Enabled = estado;
            TxtChofer.Enabled = estado;
            TxtCapaCaja.Enabled = estado;
        }

        //funcion txt
        private void imprimir()
        {
            using (StreamWriter sw = new StreamWriter(@"C:\Reportes\pedido_exp.txt"))
            {
                thisConnection.Open();
                //DateTime dt = DateTime.Now;
                sw.WriteLine("\t\t\t\t\tComercializadora GAB, S.A. de C.V.\t\t\t\t Copia1.- Ventas");
                sw.WriteLine("\t\t\t\tORDEN DE VENTA EXPORTACION/REPORTE DE EMBARQUES\t\t\t Copia2.- Bodega");
                sw.WriteLine("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t Copia3.- Embarques\r\n");
                //sw.WriteLine("\t\t\t\t\t\t\t\t\t\t {0}", codbar.Text);
                sw.WriteLine("\tCLIENTE: {0} {1} FOLIO: {2}\t FECHA: {3}", txtcliente.Text, CBCliente.SelectedItem.ToString(), txtfolio.Text, txtfecha.Text);
                sw.WriteLine("\tTRANSPORTISTA: {0} CARTA PORTE: {1} \t\t\t IRAPUATO GTO.\r\n", CBTransportista.SelectedItem.ToString().Trim(), txttalonemb.Text.Trim());

                sw.WriteLine("\tDETALLE DE LA ORDEN DE VENTA\r\n");
                sw.WriteLine("\tLinea \t Producto \t\t\t\t\t\t\t\t\t\t Pedido \t Tarima \t\t\t\t\t\t\t Tar. Aprox. \tObs\r\n");
                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select prod_num_tarimas from tb_cat_producto where prod_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "'";
                    reader2 = cmnd2.ExecuteReader();
                    while (reader2.Read())
                    {
                        prod_num_tarimas = reader2.GetSqlDecimal(0).Value;
                    }
                    reader2.Close();
                    if (prod_num_tarimas == 0)
                        resulta = 0;

                    else
                        //resulta = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) / 1;
                        resulta = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) / prod_num_tarimas;

                    sw.WriteLine("\t{0} {1} {2}|     {3,3}    |\t{4,10}\t|\t{5,5}| {6,30}|", Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value), Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value), Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value), Convert.ToString(DGV2.Rows[i].Cells["nombre_tarima"].Value), resulta.ToString("0.00"), Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value));
                    //sw.WriteLine("\t\t\t\t\t\t\t\t\t\t\t\t\t   |         \t|\t{0,10}\t\t\t\t\t|\t\t|\t\t\t\t\t\t |", Convert.ToString(DGV2.Rows[i].Cells["subcliente"].Value));
                    sw.WriteLine("\t{0,21}     {1,20} {2,20}|     {3,3}    |\t{4,34}\t\t\t\t\t\t|\t{5,5}| {6,30} |", FechaCadTexto(DGV2.Rows[i]), "", "", "", Convert.ToString(DGV2.Rows[i].Cells["subcliente"].Value), "", "");
                    //sw.WriteLine("\t Pedimento: {0,15} Fecha: {1,10}\t Aduana: {2,5}", Convert.ToString(DGV2.Rows[i].Cells["pedimento"].Value), Convert.ToString(DGV2.Rows[i].Cells["fecha_ped"].Value), Convert.ToString(DGV2.Rows[i].Cells["aduana"].Value));                                        
                    sumapedido = sumapedido + Convert.ToInt32(DGV2.Rows[i].Cells["cajas"].Value);
                    sumataraprox = sumataraprox + resulta;
                    enteros = Math.Truncate(resulta);
                    sumacomp = sumacomp + enteros;
                    decimales = enteros - resulta;
                    sumasplit = sumasplit + decimales;
                }

                sw.WriteLine("\t\t\t\t\t\t\t\t\t\t\t\t\t\t   {0}\t\t Tarimas Aprox: {1}", sumapedido, sumataraprox.ToString("0.00"));
                sw.WriteLine("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t Completas: {0}", sumacomp);
                sw.WriteLine("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t Split: {0}\r\n\r\n\r\n", (sumasplit * -1).ToString("0.00"));
                sw.WriteLine("\t{0}\r\n", txtobs.Text);
                sw.WriteLine("\tPONER TARIMAS INTERMEDIAS DE SOPORTE Y AJUSTAR LAS 5 GATAS");
                sw.WriteLine("\t" + leyenda1);
                sw.WriteLine("\tPedido:{0}\t\t\t Ryan digital: {1}\r\n", txtpedsigma.Text, ryandigital);
                sw.WriteLine("\tCargo: __________ \t Destino: ___________________ \tSello: {0}\r\n", txtsello.Text);
                //sw.WriteLine("\tMe comprometo a mantener la temperautia a: ________F Nombre chofer: {0}\r\n", txtnomtrans.Text);
                sw.WriteLine("\tFirma chofer: _____________________ TERMOGRAFO:_________________ \r\n");
                sw.WriteLine("\tEmpezo a cargar camion (horas): __________  Termino (horas): __________");
                sw.Close();
                sumataraprox = sumacomp = sumasplit = decimales = enteros = 0;
                sumapedido = 0;
                thisConnection.Close();
                //System.Diagnostics.Process.Start(@"C:\\Reportes\pedido_exp.txt");
                //System.Diagnostics.Process.Start("wordpad.exe", @"C:\\Reportes\pedido_exp.txt");
                // SE CAMBIO A NOTEPAD PORQUE WINDOWS SE ACTUALIZO Y BORRO LA APLIACION 14 DIC 2024
                System.Diagnostics.Process.Start("notepad.exe", @"C:\\Reportes\pedido_exp.txt");
            }
        }

        //método que limpia todo
        private void limpiarTextBoxes(Control parent)
        {
            //Limpiar de manera rapida
            foreach (Control c in parent.Controls)
            {
                if (c is TextBox)
                {
                    TextBox txt = (TextBox)c;
                    txt.Text = "";
                    txt.Enabled = false;
                }
                if (c.Controls.Count > 0)
                {
                    limpiarTextBoxes(c);
                }
            }

            DataTable DT1 = (DataTable)DGV2.DataSource;
            if (DT1 != null)
                DT1.Clear();


            btnAlta.Enabled = true;
            btnConsulta.Enabled = true;
            btnGuardar.Enabled = false;
            btnCancel.Enabled = false;
            CBmoneda.Enabled = false;
            CBTransportista.Enabled = false;
            CBProducto.Enabled = false;
            CBCliente.Enabled = false;
            CBmoneda.SelectedIndex = -1;
            CBTransportista.SelectedIndex = -1;
            CBProducto.SelectedIndex = -1;
            CBCliente.SelectedIndex = -1;
            cbdestino.SelectedIndex = -1;
            cbdestino.Items.Clear();
            txtmn.Enabled = true;
            txtimporte.Text = "0.00";
            txtcompleped.Text = "0";
            LblElaboro.Text = "";
        }

        #region OBTENER EL TIPO DEL PRODUCTO 
        public string ObtenerProdTipo(string prod_clave)
        {
            string prodTipo = "";

            SqlCommand cmd = null;
            SqlDataReader reader = null;

            try
            {
                cmd = thisConnection.CreateCommand();
                cmd.CommandText = @"
            SELECT prod_tipo
            FROM dbo.tb_cat_producto
            WHERE prod_clave = @prod_clave
        ";

                cmd.Parameters.AddWithValue("@prod_clave", prod_clave);

                reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    prodTipo = reader["prod_tipo"].ToString().Trim();
                }
            }
            catch
            {
                // legacy: silencioso, no toca conexión
                prodTipo = "";
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                    reader.Dispose();
                }

                if (cmd != null)
                {
                    cmd.Dispose();
                }
            }

            return prodTipo;
        }

        #endregion

        //botón alta
        private void btnAlta_Click(object sender, EventArgs e)
        {
            try
            {
                btnAlta.Enabled = false;
                thisConnection.Open();
                txtfecha.Text = DateTime.Now.ToShortDateString();
                habilita(true);
                CBmoneda.SelectedIndex = 1;
                opcion = 1;
                CBmoneda_SelectionChangeCommitted(sender, e);
                cmnd2 = thisConnection.CreateCommand();
                //cmnd2 = new SqlCommand("SELECT IDENT_CURRENT('tb_mstr_pedidos_exp')", thisConnection);
                cmnd2 = new SqlCommand("SELECT max(pdn_folio) from tb_mstr_pedidos_exp", thisConnection);
                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    sig = reader2.GetSqlValue(0).ToString();
                }
                sige = Convert.ToInt32(sig) + 1;
                txtfolio.Text = sige.ToString();
                reader2.Close();
                thisConnection.Close();
                btnConsulta.Enabled = false;
                btnActu.Enabled = false;
                btnbaja.Enabled = false;
                btnGuardar.Enabled = true;
                btnCancel.Enabled = true;
                btncopiarpedido.Enabled = true;
                BtnPedTom.Enabled = true;
                DGV2.Enabled = true;
                DGV2.ReadOnly = false;
                DGV2.Columns[0].ReadOnly = true;
                DGV2.Columns[1].ReadOnly = true;
                DGV2.Columns[2].ReadOnly = true;
                CBmonfle.SelectedIndex = 1;
                txtcliente.Focus();
                DtHrCita.Text = "00:00";
                CBtipoflete.Text = "TERRESTRE";
                CBCliente.Focus();
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void OrdVenExp_Load(object sender, EventArgs e)
        {
            #region VALIDA Y FORZA ACTUALIZAR
            if (ValidaActualizacion())
            {
                requiereActualizar = true;
                this.Close(); // Esto disparará el flujo de cierre inmediatamente
            }
            #endregion
            //LblVer.Text = "Ver: " + File.GetLastWriteTime(@"c:\sisgabweb\ConocimientoEmbarque.exe").ToString();
            MakeDataTabletmgrdhis();
            try
            {
                CreaTabla();
                Faltantes.Columns.Add("ClavePt", typeof(string));
                Faltantes.Columns.Add("Nomprod", typeof(string));
                Faltantes.Columns.Add("Cajas", typeof(decimal));
                Faltantes.Columns.Add("Importe", typeof(decimal));

                thisConnection.Open();

                #region combo box clientes
                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "select distinct cnte_nombre from tb_cat_cliente order by cnte_nombre";
                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    CBCliente.Items.Add(reader2.GetSqlString(0).ToString());
                }
                reader2.Close();
                #endregion

                transportista();

                //string query = "select cnte_clave, entr_clave, entr_nombre, Rtrim(entr_NOMBRE) + ' ' + RTRIM(ENTR_CIUDAD) AS NOMBRE  from tb_cat_entregas order by entr_nombre";
                //da = new SqlDataAdapter(query, thisConnection);
                //da.Fill(ds1, "sub");
                //subcnte = ds1.Tables["sub"];

                string query = "SELECT usu_login, usu_nombre FROM tb_cat_usuarios order by usu_login";
                da = new SqlDataAdapter(query, thisConnection);
                da.Fill(ds1, "User");
                CatUser = ds1.Tables["User"];
                //query = "select emp_clave, emp_nombre from tb_cat_empaques where emp_nombre like '%tarima%' order by emp_nombre";
                query = "select id_tarima as emp_clave, nom_tarima as emp_nombre from tb_cat_tarima order by nom_tarima";
                da = new SqlDataAdapter(query, thisConnection);
                da.Fill(ds1, "tar");
                tarimas = ds1.Tables["tar"];
                string cadena = "Select uerp from tb_mstr_ExisteUERP";
                SqlCommand cmd1 = new SqlCommand(cadena, thisConnection);
                ExisteUERP = Convert.ToString(cmd1.ExecuteScalar());
                thisConnection.Close();

                foreach (DataRow dr in tarimas.Rows)
                {
                    nombre_tarima.Items.Add(Convert.ToString(dr["emp_nombre"].ToString()).Trim());
                }


                var inv_pt = "C:\\Reportes\\pedido_exp.txt";

                if (!File.Exists(inv_pt))
                {
                    var fileStream = File.Create(inv_pt);
                }
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }
        #region METODO PARA VALIDAR LA ACTUALIZACION DEL EJECUTABLE
        public bool ValidaActualizacion()
        {
            string rutaServerTxt = @"\\gabira1\sisgabweb\Valida.txt";
            string rutaServerExe = @"\\gabira1\sisgabweb\OrdenVentaExp.exe";
            string rutaLocalExe = @"c:\sisgabweb\OrdenVentaExp.exe";

            try
            {
                // 1. Validamos que el archivo de control en el servidor exista
                if (File.Exists(rutaServerTxt) && File.Exists(rutaServerExe) && File.Exists(rutaLocalExe))
                {
                    DateTime fechaLocal = File.GetLastWriteTime(rutaLocalExe);
                    DateTime fechaServer = File.GetLastWriteTime(rutaServerExe);

                    // 2. Comparamos fechas
                    if (fechaServer > fechaLocal)
                    {
                        MessageBox.Show("Hay una versión más reciente. El sistema se cerrará para actualizarse.\n\nPor favor, vuelva a abrir el programa cuando finalice.",
                                        "Actualización Disponible",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);
                        return true;
                    }
                }
            }
            catch (IOException ex)
            {
                // Si la red se cae o el archivo está bloqueado, atrapamos el error para que no truene la app
                Console.WriteLine($"Error al validar actualización: {ex.Message}");
            }

            return false;
        }
        #endregion

        //combobox para el tipo de moneda        
        private void CBmoneda_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (CBmoneda.SelectedIndex == 0)
            {
                DGV2.Columns[0].HeaderText = "";
                DGV2.Columns["precios"].HeaderText = "Precio M.N.";//7
                DGV2.Columns["total"].HeaderText = "Total M.N.";    //8             
                moneda = "PESOS";
                txtmn.Text = "0";
                //DGV2.Enabled = true;
            }
            if (CBmoneda.SelectedIndex == 1)
            {
                DGV2.Columns["precios"].HeaderText = "Precio USD";
                DGV2.Columns["total"].HeaderText = "Total USD";
                moneda = "USD";
                txtusd.Text = "0";
                //DGV2.Enabled = true;
            }
            if (CBmoneda.SelectedIndex == 2)
            {
                DGV2.Columns["precios"].HeaderText = "Precio EUROS";
                DGV2.Columns["total"].HeaderText = "Total EUROS";
                moneda = "EUROs";
                txteuro.Text = "0";
                //DGV2.Enabled = true;
            }
        }

        //combobox del cliente
        private void CBCliente_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                thisConnection.Open();
                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "select distinct cnte_clave from tb_cat_cliente where cnte_nombre = '" + CBCliente.SelectedItem.ToString() + "'";
                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    txtcliente.Text = reader2.GetSqlString(0).ToString();
                }
                reader2.Close();
                if ((txtcliente.Text.Trim() == "GABIN") || (txtcliente.Text.Trim() == "GAB2") || (txtcliente.Text.Trim() == "GAB3") ||
                    (txtcliente.Text.Trim() == "GAB4") || (txtcliente.Text.Trim() == "GAB5") || (txtcliente.Text.Trim() == "MASTR") ||
                    (txtcliente.Text.Trim() == "MAST2") || (txtcliente.Text.Trim() == "GABOP"))
                {
                    cbpagoGAB.Checked = true;
                    cbpagoGAB.Enabled = false;
                }
                else
                    cbpagoGAB.Enabled = true;
                txtcliente.Text = txtcliente.Text.ToUpper();
                cat_productos();
                subcliente.Items.Clear();
                subcnte.Rows.Clear();
                cbdestino.Items.Clear();
                string query = "select cnte_clave, entr_clave, entr_nombre, Rtrim(entr_NOMBRE) + ' ' + RTRIM(ENTR_CIUDAD) AS NOMBRE,ENTR_DIRECCION,ENTR_CIUDAD,ENTR_ESTADO,ENTR_CP,ENTR_COLONIA,ENTR_RFC  from tb_cat_entregas " +
                                   " where cnte_clave = '" + txtcliente.Text + "' order by entr_nombre";
                da = new SqlDataAdapter(query, thisConnection);
                da.Fill(ds1, "sub");
                subcnte = ds1.Tables["sub"];
                foreach (DataRow dr in subcnte.Rows)  //.Select("cnte_clave = '" + txtcliente.Text + "'"))
                //foreach (DataRow dr in subcnte.Select("cnte_clave = '" + txtcliente.Text + "'"))
                {
                    subcliente.Items.Add(dr["nombre"].ToString().Trim()); //entr_nombre
                    cbdestino.Items.Add(dr["entr_nombre"].ToString().Trim());
                }
                thisConnection.Close();
                CBProducto.Enabled = true;
                txtclaveprod.Enabled = true;
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        //textbox clave del cliente
        private void txtcliente_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                try
                {
                    txtcliente.Text = txtcliente.Text.ToUpper();
                    thisConnection.Open();
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select distinct cnte_nombre from tb_cat_cliente where cnte_clave = '" + txtcliente.Text + "'";
                    reader2 = cmnd2.ExecuteReader();
                    while (reader2.Read())
                    {
                        CBCliente.SelectedItem = reader2.GetSqlString(0).ToString();
                    }
                    if (reader2.HasRows == false)
                    {
                        reader2.Dispose();
                        thisConnection.Close();
                        MessageBox.Show("No existe ningún cliente con esa clave, favor de verificarla", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        txtcliente.Focus();
                        txtcliente.SelectAll();
                        return;
                    }
                    if ((txtcliente.Text.Trim() == "GABIN") || (txtcliente.Text.Trim() == "GAB2") || (txtcliente.Text.Trim() == "GAB3") ||
                         (txtcliente.Text.Trim() == "GAB4") || (txtcliente.Text.Trim() == "GAB5") || (txtcliente.Text.Trim() == "MASTR") ||
                         (txtcliente.Text.Trim() == "MAST2") || (txtcliente.Text.Trim() == "GABOP"))
                    {
                        cbpagoGAB.Checked = true;
                        cbpagoGAB.Enabled = false;
                    }
                    else
                        cbpagoGAB.Enabled = true;

                    reader2.Dispose();
                    cat_productos();
                    subcliente.Items.Clear();
                    cbdestino.Items.Clear();
                    subcnte.Rows.Clear();
                    string query = "select cnte_clave, entr_clave, entr_nombre, Rtrim(entr_NOMBRE) + ' ' + RTRIM(ENTR_CIUDAD) AS NOMBRE,ENTR_DIRECCION,ENTR_CIUDAD,ENTR_ESTADO,ENTR_CP,ENTR_COLONIA, ENTR_RFC  from tb_cat_entregas " +
                                   " where cnte_clave = '" + txtcliente.Text + "' order by entr_nombre";
                    da = new SqlDataAdapter(query, thisConnection);
                    da.Fill(ds1, "sub");
                    subcnte = ds1.Tables["sub"];
                    foreach (DataRow dr in subcnte.Rows)  //.Select("cnte_clave = '" + txtcliente.Text + "'"))
                    {
                        subcliente.Items.Add(dr["nombre"].ToString().Trim()); //entr_nombre
                        cbdestino.Items.Add(dr["entr_nombre"].ToString().Trim());
                    }
                    thisConnection.Close();
                    CBProducto.Enabled = true;
                    txtclaveprod.Enabled = true;
                    CBProducto.Focus();
                }
                catch (SqlException ex)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                    MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                catch (Exception ex1)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                    MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }

        //combobox del producto
        private void CBProducto_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                thisConnection.Open();
                txtcliente.Text = txtcliente.Text.ToUpper();
                string prodname = CBProducto.SelectedItem.ToString().Trim();
                prodname = prodname.Replace("'", "''");
                string tar = "", nom_tar = "";
                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "select prod_clave, lin_clave, prod_tarimas, prod_clave_gabinc, prod_nomb_ingles from tb_cat_producto where prod_nombre = '" + prodname + "'";
                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    txtclaveprod.Text = reader2.GetSqlString(0).ToString();
                    lin_clave = reader2.GetSqlString(1).ToString();
                    tar = reader2.GetValue(2).ToString().Trim();
                    item_number = reader2.GetValue(3).ToString().Trim();
                    TA = reader2.GetValue(4).ToString().Substring(0, 49);
                    TB = reader2.GetValue(4).ToString().Substring(50, 49);
                    TC = reader2.GetValue(4).ToString().Substring(100, 49);
                    TD = reader2.GetValue(4).ToString().Substring(150, 49);
                    TE = reader2.GetValue(4).ToString().Substring(200, 49);
                    if (txtcliente.Text.Trim() == "GABIN" || txtcliente.Text.Trim() == "GAB2" || txtcliente.Text.Trim() == "GAB3" || txtcliente.Text.Trim() == "GAB4" || (txtcliente.Text.Trim() == "GABOP"))
                    {
                        if (item_number.Trim() == "")
                        {
                            MessageBox.Show("No se puede agregar el producto ya que no tiene número de item, favor de capturarlo en el catálogo de productos", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            thisConnection.Close();
                            return;
                        }
                    }
                }
                reader2.Dispose();


                #region cuenta el numero de filas del DGV2 y solo permite 35 productos
                if (DGV2.Rows.Count >= 0)
                {
                    if (DGV2.Rows.Count == 35)
                    {
                        MessageBox.Show("No se pueden agregar mas productos a la orden", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        thisConnection.Close();
                        return;
                    }
                    DataRow encrowc = ds1.Tables["tmgrdped"].Rows.Find(txtclaveprod.Text);
                    if (encrowc != null)
                    {
                        MessageBox.Show("Este componente ya ha sido agregado");
                        thisConnection.Close();
                        return;
                    }
                }

                #endregion

                #region verifica la existencia del producto en el ERP
                //07/10/2022

                if (ExisteUERP != "N")
                {
                    thisConnection2.Open();
                    cmnd2 = thisConnection2.CreateCommand();
                    cmnd2.CommandText = "select ArtCod, ArtNom from articulos where ArtCod = '" + txtclaveprod.Text + "'";
                    reader2 = cmnd2.ExecuteReader();
                    while (reader2.Read())
                    {
                        ArtCod = reader2.GetSqlString(0).ToString();
                    }
                    reader2.Close();
                    thisConnection2.Close();
                    if (ArtCod.Trim() != txtclaveprod.Text.Trim())
                    {
                        MessageBox.Show("El código del producto " + ArtCod.Trim() + " " + CBProducto.SelectedItem.ToString() + "\r\n" + "No esta dado de alta en el sistema USABIAGA ERP\r\n" + "Verificar con compras la alta del producto, gracias \r\n" + "No se puede generar el pedido" + "Alta de pedidos BCN", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        thisConnection.Close();
                        return;
                    }
                    //fin 07/10/2022
                }
                #endregion

                cmnd2 = thisConnection.CreateCommand();
                //cmnd2.CommandText = "SELECT tdp.pdn_precio_usd, tmpn.pdn_folio FROM tb_mstr_pedidos_nal tmpn, tb_det_pedidos tdp WHERE tmpn.cnte_clave='" + txtcliente.Text + "' AND tdp.pdn_folio = tmpn.pdn_folio AND tdp.prod_clave='" + txtclaveprod.Text + "' ORDER BY tmpn.pdn_fecha DESC";
                if (CBmoneda.SelectedIndex == 0)
                    cmnd2.CommandText = "SELECT tdp.pdn_precio_mn FROM tb_mstr_pedidos_nal tmpn, tb_det_pedidos tdp WHERE tmpn.cnte_clave='" + txtcliente.Text + "'  AND tdp.prod_clave='" + txtclaveprod.Text + "' ORDER BY tmpn.pdn_fecha DESC";

                if (CBmoneda.SelectedIndex == 1 || CBmoneda.SelectedIndex == 2)
                    cmnd2.CommandText = "SELECT tdp.pdn_precio_usd FROM tb_mstr_pedidos_nal tmpn, tb_det_pedidos tdp WHERE tmpn.cnte_clave='" + txtcliente.Text + "'  AND tdp.prod_clave='" + txtclaveprod.Text + "' ORDER BY tmpn.pdn_fecha DESC";

                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    precio = reader2.GetSqlDecimal(0).Value;
                    encontro = true;
                    break;
                }
                reader2.Close();
                if (encontro == false)
                {
                    cmnd4 = thisConnection.CreateCommand();
                    cmnd4.CommandText = "select prod_precio_usd from tb_cat_producto where prod_clave='" + txtclaveprod.Text + "'";
                    reader4 = cmnd4.ExecuteReader();
                    while (reader4.Read())
                    {
                        precio = reader4.GetSqlDecimal(0).Value;
                    }
                    reader4.Close();
                }
                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "select emp_nombre from tb_cat_empaques where emp_clave = '" + tar + "'";
                nom_tar = Convert.ToString(cmnd2.ExecuteScalar()).Trim();
                thisConnection.Close();
                encontro = false;

                DataRow drc = tmgrdped.NewRow();
                drc["linea"] = lin_clave;
                drc["producto"] = txtclaveprod.Text;
                drc["nombre"] = CBProducto.SelectedItem.ToString(); ;
                drc["precio"] = precio;
                drc["cantidad"] = 0;
                drc["clave_tarima"] = tar;
                drc["nom_tarima"] = nom_tar;
                drc["subcliente"] = "";
                drc["subcnte_clave"] = "";
                drc["po_number"] = "";
                drc["total"] = 0;
                drc["observaciones"] = " ";
                drc["transito"] = "";
                drc["especial"] = "Obs Especial";
                drc["special"] = "";
                tmgrdped.Rows.Add(drc);
                DGV2.DataSource = tmgrdped;

                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    this.DGV2.Rows[i].HeaderCell.Value = (i + 1).ToString();
                }
                DGV2.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader);
                //DGV2.ReadOnly = false;
                DGV2.Focus();
                DGV2.CurrentCell = DGV2.Rows[DGV2.Rows.Count - 1].Cells["cajas"];
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "sisgab", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "sisgab", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        //textbox clave del producto
        private void txtclaveprod_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                try
                {
                    string tar = "", nom_tar = "";
                    thisConnection.Open();
                    txtcliente.Text = txtcliente.Text.ToUpper();
                    txtclaveprod.Text = txtclaveprod.Text.ToUpper();
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select prod_nombre, lin_clave,prod_clave_gabinc, prod_nomb_ingles from tb_cat_producto where prod_clave = '" + txtclaveprod.Text + "'";
                    reader2 = cmnd2.ExecuteReader();
                    while (reader2.Read())
                    {
                        CBProducto.SelectedItem = reader2.GetSqlString(0).ToString().Trim();
                        lin_clave = reader2.GetSqlString(1).ToString();
                        item_number = reader2.GetValue(2).ToString().Trim();
                        TA = reader2.GetValue(3).ToString().Substring(0, 49);
                        TB = reader2.GetValue(3).ToString().Substring(50, 49);
                        TC = reader2.GetValue(3).ToString().Substring(100, 49);
                        TD = reader2.GetValue(3).ToString().Substring(150, 49);
                        TE = reader2.GetValue(3).ToString().Substring(200, 49);
                        if (txtcliente.Text.Trim() == "GABIN" || txtcliente.Text.Trim() == "GAB2" || txtcliente.Text.Trim() == "GAB3" || txtcliente.Text.Trim() == "GAB4" || (txtcliente.Text.Trim() == "GABOP"))
                        {
                            if (item_number.Trim() == "")
                            {
                                MessageBox.Show("No se puede agregar el producto ya que no tiene número de item, favor de capturarlo en el catálogo de productos", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                thisConnection.Close();
                                return;
                            }
                        }
                    }
                    if (reader2.HasRows == false || CBProducto.SelectedIndex == -1)
                    {
                        thisConnection.Close();
                        MessageBox.Show("Producto no encontrado, favor de verificar la clave", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    reader2.Close();

                    #region cuenta el numero de filas del DGV2 y solo permite 35 productos
                    if (DGV2.Rows.Count >= 0)
                    {
                        if (DGV2.Rows.Count == 35)
                        {
                            MessageBox.Show("No se pueden agregar mas productos a la orden", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            thisConnection.Close();
                            return;
                        }
                        DataRow encrowc = ds1.Tables["tmgrdped"].Rows.Find(txtclaveprod.Text);
                        if (encrowc != null)
                        {
                            MessageBox.Show("Este componente ya ha sido agregado");
                            thisConnection.Close();
                            return;
                        }
                    }

                    #endregion

                    # region verifica la existencia del producto en el ERP
                    //07/10/2022

                    if (ExisteUERP != "N")
                    {
                        thisConnection2.Open();
                        cmnd2 = thisConnection2.CreateCommand();
                        cmnd2.CommandText = "select ArtCod, ArtNom from articulos where ArtCod = '" + txtclaveprod.Text + "'";
                        reader2 = cmnd2.ExecuteReader();
                        while (reader2.Read())
                        {
                            ArtCod = reader2.GetSqlString(0).ToString();
                        }
                        reader2.Close();
                        thisConnection2.Close();
                        if (ArtCod.Trim() != txtclaveprod.Text.Trim())
                        {
                            MessageBox.Show("El código del producto " + ArtCod.Trim() + " " + CBProducto.SelectedItem.ToString() + "\r\n" + "No esta dado de alta en el sistema USABIAGA ERP\r\n" + "Verificar con compras la alta del producto, gracias \r\n" + "No se puede generar el pedido" + "Alta de pedidos BCN", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            thisConnection.Close();
                            return;
                        }
                        //fin 07/10/2022
                    }
                    #endregion

                    cmnd2 = thisConnection.CreateCommand();
                    if (CBmoneda.SelectedIndex == 0)
                        cmnd2.CommandText = "SELECT tdp.pdn_precio_mn FROM tb_mstr_pedidos_nal tmpn, tb_det_pedidos tdp WHERE tmpn.cnte_clave='" + txtcliente.Text + "'  AND " +
                                            "tdp.prod_clave='" + txtclaveprod.Text + "' ORDER BY tmpn.pdn_fecha DESC";

                    if (CBmoneda.SelectedIndex == 1 || CBmoneda.SelectedIndex == 2)
                        cmnd2.CommandText = "SELECT tdp.pdn_precio_usd FROM tb_mstr_pedidos_nal tmpn, tb_det_pedidos tdp WHERE tmpn.cnte_clave='" + txtcliente.Text + "'  AND " +
                                            "tdp.prod_clave='" + txtclaveprod.Text + "' ORDER BY tmpn.pdn_fecha DESC";

                    reader2 = cmnd2.ExecuteReader();
                    while (reader2.Read())
                    {
                        precio = reader2.GetSqlDecimal(0).Value;
                        encontro = true;
                        break;
                    }
                    reader2.Close();
                    if (encontro == false)
                    {
                        cmnd4 = thisConnection.CreateCommand();
                        cmnd4.CommandText = "select prod_precio_usd from tb_cat_producto where prod_clave='" + txtclaveprod.Text + "'";
                        reader4 = cmnd4.ExecuteReader();
                        while (reader4.Read())
                        {
                            precio = reader4.GetSqlDecimal(0).Value;
                        }
                        reader4.Close();
                    }

                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select emp_nombre from tb_cat_empaques where emp_clave = '" + tar + "'";
                    nom_tar = Convert.ToString(cmnd2.ExecuteScalar()).Trim();
                    thisConnection.Close();
                    encontro = false;

                    DataRow drc = tmgrdped.NewRow();
                    drc["linea"] = lin_clave;
                    drc["producto"] = txtclaveprod.Text;
                    drc["nombre"] = CBProducto.SelectedItem.ToString(); ;
                    drc["precio"] = precio;
                    drc["cantidad"] = 0;
                    drc["clave_tarima"] = tar;
                    drc["nom_tarima"] = nom_tar;
                    drc["subcliente"] = "";
                    drc["subcnte_clave"] = "";
                    drc["po_number"] = "";
                    drc["total"] = 0;
                    drc["observaciones"] = " ";
                    drc["transito"] = "";
                    drc["especial"] = "Obs Especial";
                    drc["special"] = "";
                    tmgrdped.Rows.Add(drc);
                    DGV2.DataSource = tmgrdped;

                    for (int i = 0; i < DGV2.Rows.Count; i++)
                    {
                        this.DGV2.Rows[i].HeaderCell.Value = (i + 1).ToString();
                    }
                    DGV2.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader);
                    DGV2.Enabled = true;
                    DGV2.Focus();
                    DGV2.CurrentCell = DGV2.Rows[DGV2.Rows.Count - 1].Cells["cajas"];
                }
                catch (SqlException ex)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "sisgab", ex.ToString());
                    MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                catch (Exception ex1)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "sisgab", ex1.ToString());
                    MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }

        //combobox nombre transportista
        private void CBTransportista_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                thisConnection.Open();
                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "select prov_clave from tb_cat_proveedor where prov_nombre = '" + CBTransportista.SelectedItem.ToString() + "'";
                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    txtclavetrans.Text = reader2.GetSqlString(0).ToString();
                }
                thisConnection.Close();
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void OrdVenExp_FormClosed(object sender, FormClosedEventArgs e)
        {
            #region DESCARGA LA NUEVA VERSION DEL EJECUTABLE
            //MessageBox.Show("Despues de Cerrar");
            string ArchDescarga = @"c:\sisgabweb\DownFile.exe";
            if (Actualizar == "S" && File.Exists(ArchDescarga))
            {
                string ArchNuevo = @"\\gabira1\sisgabweb\OrdenVentaExp.exe";
                if (File.Exists(ArchNuevo))
                    //File.Copy(ArchNuevo, ArchViejo, true);
                    System.Diagnostics.Process.Start(@"c:\sisgabweb\DownFile.exe", "OrdenVentaExp.exe");
            }
            #endregion
        }

        //textbox clave transportista
        private void txtclavetrans_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                try
                {
                    thisConnection.Open();
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select prov_nombre from tb_cat_proveedor where prov_clave ='" + txtclavetrans.Text + "'";
                    reader2 = cmnd2.ExecuteReader();
                    while (reader2.Read())
                    {
                        CBTransportista.SelectedItem = reader2.GetSqlString(0).ToString();
                    }
                    thisConnection.Close();
                    if (CBTransportista.SelectedIndex == -1)
                    {
                        MessageBox.Show("No existe transportista con esa clave", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (SqlException ex)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                    MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                catch (Exception ex1)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                    MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }

        ComboBox cntesub = new System.Windows.Forms.ComboBox();
        ComboBox tarim = new ComboBox();
        //evento que valida la celda del grid para que solo acepte numeros
        private void DGV2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                DataGridViewTextBoxEditingControl tb = e.Control as DataGridViewTextBoxEditingControl;
                tb.KeyPress -= DGV2_KeyPress;
                tb.KeyPress += new KeyPressEventHandler(DGV2_KeyPress);
            }
            if (DGV2.CurrentCell.ColumnIndex == 0 || DGV2.CurrentCell.ColumnIndex == 2)
            {
                txt1 = e.Control as TextBox;
                if (txt1 != null)
                {
                    txt1.KeyPress -= new KeyPressEventHandler(DGV2_KeyPress);
                    txt1.KeyPress += new KeyPressEventHandler(DGV2_KeyPress);
                }
            }
            if (DGV2.CurrentCell.ColumnIndex == 9)
            {
                cntesub = e.Control as ComboBox;
                if (cntesub != null)
                {
                    cntesub.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
                    cntesub.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
                }
                //}
            }
            if (DGV2.CurrentCell.ColumnIndex == 7)
            {
                DataGridViewComboBoxEditingControl tarim = e.Control as DataGridViewComboBoxEditingControl;
                if (tarim != null)
                {
                    tarim.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
                    tarim.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
                }
            }
        }

        //evento que evalua las teclas presionadas
        private void DGV2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.ToString() == e.KeyChar.ToString().ToUpper())
            {
                return;
            }
            e.Handled = true;
            SendKeys.Send(e.KeyChar.ToString().ToUpper());
            if (DGV2.CurrentCell.ColumnIndex == 0 || DGV2.CurrentCell.ColumnIndex == 2)
            {
                if (e.KeyChar == 8)
                {
                    e.Handled = false;
                    return;
                }

                bool IsDec = false;
                int nroDec = 0;

                for (int i = 0; i < txt1.Text.Length; i++)
                {
                    if (txt1.Text[i] == '.')
                        IsDec = true;

                    if (IsDec && nroDec++ >= 2)
                    {
                        e.Handled = true;
                        return;
                    }
                }

                if (e.KeyChar >= 48 && e.KeyChar <= 57)
                    e.Handled = false;
                else if (e.KeyChar == 46 || e.KeyChar == 47 || e.KeyChar == (char)Keys.Delete)
                    e.Handled = (IsDec) ? true : false;
                else
                    e.Handled = true;
            }
        }

        //evento que evalua 
        private void DGV2_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
            resulta = 0;
            if (DGV2.Columns[e.ColumnIndex].Name == "cajas" || DGV2.Columns[e.ColumnIndex].Name == "precios")
            {
                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    if (Convert.ToString(DGV2.Rows[i].Cells["cajas"].Value).Trim() == "")
                        DGV2.Rows[i].Cells["cajas"].Value = 0;

                    if (Convert.ToString(DGV2.Rows[i].Cells["precios"].Value).Trim() == "")
                        DGV2.Rows[i].Cells["precios"].Value = 0;
                    //if ((Convert.ToString(DGV2.Rows[i].Cells["cajas"].Value).Trim() != "") && (Convert.ToString(DGV2.Rows[i].Cells["precios"].Value).Trim() != ""))
                    //{
                    if (CBmoneda.SelectedIndex == 0)
                    {
                        DGV2.Rows[i].Cells["total"].Value = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) * Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value);
                        resulta = resulta + Convert.ToDecimal(DGV2.Rows[i].Cells["total"].Value);
                        txtmn.Text = resulta.ToString("###,##0.###");
                        txtusd.Text = string.Empty;
                        txteuro.Text = string.Empty;
                    }
                    if (CBmoneda.SelectedIndex == 1)
                    {
                        DGV2.Rows[i].Cells["total"].Value = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) * Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value);
                        resulta = resulta + Convert.ToDecimal(DGV2.Rows[i].Cells["total"].Value);
                        txtusd.Text = resulta.ToString("###,##0.###");
                        txtmn.Text = string.Empty;
                        txteuro.Text = string.Empty;
                    }
                    if (CBmoneda.SelectedIndex == 2)
                    {
                        DGV2.Rows[i].Cells["total"].Value = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) * Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value);
                        resulta = resulta + Convert.ToDecimal(DGV2.Rows[i].Cells["total"].Value);
                        txteuro.Text = resulta.ToString("###,##0.###");
                        txtmn.Text = string.Empty;
                        txtusd.Text = string.Empty;
                    }
                    //}
                }
            }
        }

        //valida que la fecha tengo el formato de fecha
        //fecha de caducidad de la fila para el insert: 'yyyyMMdd' o NULL
        private string FechaCadSql(DataGridViewRow row)
        {
            DateTime dtCad;
            if (DateTime.TryParseExact(Convert.ToString(row.Cells["fecha_cad"].Value).Trim(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dtCad))
                return "'" + dtCad.ToString("yyyyMMdd") + "'";
            return "NULL";
        }

        //texto "Cad: dd/MM/yyyy" para impresión; vacío si no hay fecha
        private string FechaCadTexto(DataGridViewRow row)
        {
            string f = Convert.ToString(row.Cells["fecha_cad"].Value).Trim();
            return f == "" ? "" : "Cad: " + f;
        }

        private Boolean EsFecha(String fecha)
        {
            try
            {
                DateTime.Parse(fecha);
                return true;
            }
            catch
            {
                return false;
            }
        }

        //valida que se haya introoducida una fecha en el campo de fecha de pedimento
        private void DGV2_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (DGV2.Columns[e.ColumnIndex].Name == "fecha_cad")
            {
                string txtFecha = Convert.ToString(e.FormattedValue).Trim();
                DateTime dtCad;
                if (txtFecha != "" && !DateTime.TryParseExact(txtFecha, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dtCad))
                {
                    MessageBox.Show("La fecha de caducidad debe tener el formato dd/MM/yyyy", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }
            }
            //Validamos si no es una fila nueva
            //if (!this.EsFecha(e.FormattedValue.ToString()) && DGV2.Rows[e.RowIndex].Cells["pedimento"].Value.ToString().Trim().Length > 0)
            //{
            //if (DGV2.Columns[e.ColumnIndex].Name == "fecha_ped")
            //{
            //    if (!this.EsFecha(e.FormattedValue.ToString()))
            //    {
            //        MessageBox.Show("El dato introducido no es de tipo fecha", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        DGV2.Rows[e.RowIndex].ErrorText = "El dato introducido no es de tipo fecha";
            //        e.Cancel = true;
            //    }
            //}                

            //if (DGV2.Columns[e.ColumnIndex].Name == "cajas")
            //{
            //    for (int i = 0; i < DGV2.Rows.Count; i++)
            //    {
            //        if (Convert.ToString(DGV2.Rows[i].Cells["precios"].Value) == "0.00")
            //        {
            //            MessageBox.Show("Favor de escribir el precio del producto", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);

            //            return;
            //        }
            //    }
            //}
            //}
        }

        bool find = false;
        //marca la fila en donde esta el error
        private void DGV2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //Elimina el mensaje de error de la cabecera de la fila
            //DGV2.Rows[e.RowIndex].ErrorText = String.Empty;
            if (DGV2.CurrentCell.ColumnIndex == 7)
            {
                DGV2.Rows[DGV2.CurrentRow.Index].Cells["tipo_tar"].Value = tme;
            }
            if (DGV2.CurrentCell.ColumnIndex == 8)
            {
                if (Convert.ToString(DGV2.CurrentCell.Value).Trim().Length > 0)
                {
                    find = false;
                    foreach (DataRow dr in subcnte.Select("cnte_clave ='" + txtcliente.Text + "' and entr_clave = '" + Convert.ToString(DGV2.CurrentCell.Value).Trim() + "'"))
                    {
                        DGV2.Rows[DGV2.CurrentRow.Index].Cells[9].Value = dr["entr_nombre"].ToString().Trim();
                        find = true;
                    }
                    if (find == false)
                    {
                        MessageBox.Show("No existe ningún subcliente con esa clave. Favor de verificar la clave del subcliente", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        DGV2.Rows[DGV2.CurrentRow.Index].Cells[9].Value = string.Empty;
                        DGV2.Rows[DGV2.CurrentRow.Index].Cells[8].Value = string.Empty;
                        DGV2.Rows[DGV2.CurrentRow.Index].Cells[8].Selected = true;
                        return;
                    }
                }
            }

            if (DGV2.CurrentCell.ColumnIndex == 6)
            {
                if (Convert.ToString(DGV2.CurrentCell.Value).Trim().Length > 0)
                {
                    find = false;
                    foreach (DataRow dr in tarimas.Select("emp_clave = '" + Convert.ToString(DGV2.CurrentCell.Value).Trim() + "'"))
                    {
                        DGV2.Rows[DGV2.CurrentRow.Index].Cells[7].Value = dr["emp_nombre"].ToString().Trim();
                        find = true;
                    }
                    if (find == false)
                    {
                        MessageBox.Show("No existe ninguna tarima con esa clave. Favor de verificar la clave de la tarima", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        DGV2.Rows[DGV2.CurrentRow.Index].Cells[7].Value = string.Empty;
                        DGV2.Rows[DGV2.CurrentRow.Index].Cells[6].Value = string.Empty;
                        DGV2.Rows[DGV2.CurrentRow.Index].Cells[6].Selected = true;
                        return;
                    }
                }
            }
            // se bloqueo este codigo para no bloquear la modificacion RCC 1 de mayo 2025
            if (opcion == 3) // EN ESTADO DE MODIFICACION
                if (DGV2.Columns[e.ColumnIndex].Name == "cajas" && Convert.ToInt32(DGV2.CurrentRow.Cells["Cajas"].Value) != CantMod)
                //        ValidaSplit(txtfolio.Text, DGV2.CurrentRow.Cells["clave_producto"].Value.ToString(), Convert.ToInt32(DGV2.CurrentRow.Cells["cajas"].Value), DGV2.CurrentRow.Cells["nombre"].Value.ToString()); 
                {
                    Int32 CantDif = Convert.ToInt32(DGV2.CurrentRow.Cells["Cajas"].Value) - CantMod;
                    decimal ImpDif = CantDif * Convert.ToDecimal(DGV2.CurrentRow.Cells["precios"].Value);
                    Faltantes.Rows.Add(DGV2.Rows[DGV2.CurrentRow.Index].Cells["clave_producto"].Value.ToString(), DGV2.Rows[DGV2.CurrentRow.Index].Cells["Nombre"].Value.ToString(), CantDif, ImpDif);
                }

            resulta = 0;
            if (DGV2.Columns[e.ColumnIndex].Name == "cajas" || DGV2.Columns[e.ColumnIndex].Name == "precios")
            {

                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    if (Convert.ToString(DGV2.Rows[i].Cells["cajas"].Value).Trim() == "")
                        DGV2.Rows[i].Cells["cajas"].Value = 0;

                    if (Convert.ToString(DGV2.Rows[i].Cells["precios"].Value).Trim() == "")
                        DGV2.Rows[i].Cells["precios"].Value = 0;
                    if (DGV2.Columns[e.ColumnIndex].Name == "precios")
                    {
                        if (Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) > 50)
                            MessageBox.Show("El Precio Capturado Excede los 50 Favor de Revisarlo!!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                    }
                    //if ((Convert.ToString(DGV2.Rows[i].Cells["cajas"].Value).Trim() != "") && (Convert.ToString(DGV2.Rows[i].Cells["precios"].Value).Trim() != ""))
                    //{
                    if (CBmoneda.SelectedIndex == 0)
                    {
                        DGV2.Rows[i].Cells["total"].Value = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) * Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value);
                        resulta = resulta + Convert.ToDecimal(DGV2.Rows[i].Cells["total"].Value);
                        txtmn.Text = resulta.ToString("###,##0.###");
                        txtusd.Text = string.Empty;
                        txteuro.Text = string.Empty;
                    }
                    if (CBmoneda.SelectedIndex == 1)
                    {
                        DGV2.Rows[i].Cells["total"].Value = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) * Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value);
                        resulta = resulta + Convert.ToDecimal(DGV2.Rows[i].Cells["total"].Value);
                        txtusd.Text = resulta.ToString("###,##0.###");
                        txtmn.Text = string.Empty;
                        txteuro.Text = string.Empty;
                    }
                    if (CBmoneda.SelectedIndex == 2)
                    {
                        DGV2.Rows[i].Cells["total"].Value = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) * Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value);
                        resulta = resulta + Convert.ToDecimal(DGV2.Rows[i].Cells["total"].Value);
                        txteuro.Text = resulta.ToString("###,##0.###");
                        txtmn.Text = string.Empty;
                        txtusd.Text = string.Empty;
                    }
                    //}
                }
            }
            if (DGV2.Columns[e.ColumnIndex].Name == "pedimento")
            {
                string txtnumped = Convert.ToString(DGV2.Rows[e.RowIndex].Cells["pedimento"].Value).Trim();
                if ((txtnumped.Trim() != "") || (txtnumped.Contains(' ')))
                {
                    for (int j = 0; j < Convert.ToString(DGV2.Rows[e.RowIndex].Cells["pedimento"].Value).Trim().Length; j++)
                    {
                        string pedimento = txtnumped;
                        if (j == 2)
                            txtnumped = pedimento.Insert(2, "  ");
                        if (j == 6)
                            txtnumped = pedimento.Insert(6, "  ");
                        if (j == 12)
                            txtnumped = pedimento.Insert(12, "  ");
                    }
                    DGV2.Rows[e.RowIndex].Cells["pedimento"].Value = txtnumped;
                    //}
                }
            }

        }

        //se pone por default la fecha actual
        private void DGV2_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["fecha_ped"].Value = DateTime.Now.ToShortDateString();
        }

        //evento donde se borra fila del DGV
        private void DGV2_KeyDown(object sender, KeyEventArgs e)
        {
            int XSaint = 0, fila = 0, indice = 0;
            decimal prec = 0;
            if (e.KeyCode == Keys.Delete)
            {
                string existesplit = ValidaSplitBorrado(txtfolio.Text, DGV2.Rows[DGV2.CurrentRow.Index].Cells["clave_producto"].Value.ToString(), 0, DGV2.Rows[DGV2.CurrentRow.Index].Cells["nombre"].Value.ToString());

                if (existesplit == "0")
                {
                    if (DGV2.CurrentCell.ColumnIndex == 1)
                    {
                        if (MessageBox.Show("Desea Eliminar el Producto " + System.Environment.NewLine + System.Environment.NewLine + DGV2.Rows[DGV2.CurrentRow.Index].Cells["nombre"].Value.ToString().Trim(), "Eliminar Producto", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            prec = Convert.ToDecimal(DGV2.Rows[DGV2.CurrentRow.Index].Cells["total"].Value);
                            Faltantes.Rows.Add(DGV2.Rows[DGV2.CurrentRow.Index].Cells["clave_producto"].Value.ToString(), DGV2.Rows[DGV2.CurrentRow.Index].Cells["Nombre"].Value.ToString(), Convert.ToDecimal(DGV2.Rows[DGV2.CurrentRow.Index].Cells["cajas"].Value), Convert.ToDecimal(DGV2.Rows[DGV2.CurrentRow.Index].Cells["total"].Value));
                            e.Handled = true;
                            XSaint = DGV2.CurrentRow.Index;
                            precio = Convert.ToDecimal(DGV2.Rows[DGV2.CurrentRow.Index].Cells["total"].Value);
                            foreach (DataRow row in tmgrdped.Rows)
                            {
                                if (Convert.ToString(DGV2.Rows[DGV2.CurrentRow.Index].Cells["clave_producto"].Value) == Convert.ToString(row["producto"].ToString()))
                                {
                                    indice = fila;
                                }
                                fila++;
                            }
                            tmgrdped.Rows[indice].Delete();
                            if (CBmoneda.SelectedIndex == 0)
                                txtmn.Text = (Convert.ToDecimal(txtmn.Text) - precio).ToString();
                            if (CBmoneda.SelectedIndex == 1)
                                txtusd.Text = (Convert.ToDecimal(txtusd.Text) - precio).ToString();
                            if (CBmoneda.SelectedIndex == 2)
                                txteuro.Text = (Convert.ToDecimal(txteuro.Text) - precio).ToString();
                            for (int i = 0; i < DGV2.Rows.Count; i++)
                            {
                                this.DGV2.Rows[i].HeaderCell.Value = (i + 1).ToString();
                            }
                        }
                    }
                }
            }
        }

        //botón guardar
        private void btnGuardar_Click(object sender, EventArgs e)
        {
            switch (opcion)
            {
                #region guardar
                case 1:

                    if (DtHrCita.Text.ToString().Trim() == "00:00")
                    {
                        MessageBox.Show("La hora no puede ser 00:00, Favor de poner un horario de salida real para considerar el armado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DtHrCita.Focus();
                        return;
                    }


                    if (cbdestino.SelectedIndex != -1)
                    {
                        if (cbdestino.SelectedItem.ToString().Trim() == "MCALLEN, TX.")
                        {
                            for (int i = 0; i < DGV2.Rows.Count; i++)
                            {
                                if ((Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("APIO")) || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("LECHUGA"))
                                    || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("KALE")) || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("SWISS CHARD"))
                                    || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("RED SWISS CHARD")) || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("CILANTRO"))
                                    || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("RED KALE")))
                                    leyenda1 = "ASEGURARSE QUE EN ESTE EMBARQUE HAYA UNA DISTANCIA DE 18 PULGADAS ENTRE LA PARTE ALTA DEL PALLET Y EL TECHO PARA LOS QUE CRUZAN POR " + cbdestino.SelectedItem.ToString().Trim();
                            }
                        }
                    }
                    if (cbdestino.SelectedIndex == -1)
                        destino = "";
                    else
                        destino = cbdestino.SelectedItem.ToString().Trim();

                    if (cbpagoGAB.Checked == true)
                    {
                        if (txttalonemb.Text.Trim() == "")
                        {
                            MessageBox.Show("No se ha capturado el talon de embarque", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            txttalonemb.Focus();
                            return;
                        }
                        if (txtcompleped.Text.Trim() == "0")
                        {
                            if (Convert.ToDecimal(txtimporte.Text) == 0)
                            {
                                MessageBox.Show("No se ha capturado el monto del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                txtimporte.Focus();
                                return;
                            }
                        }
                        if (Convert.ToDecimal(txtimporte.Text) > 6000)
                        {
                            MessageBox.Show("El Monto del Flete NO Puede ser Mayor a 6,000", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            txtimporte.Focus();
                            return;
                        }
                    }

                    if (txttalonemb.Text.Trim() != "")
                    {
                        if (txtcompleped.Text.Trim() == "0")
                        {
                            if (Convert.ToDecimal(txtimporte.Text) == 0)
                            {
                                MessageBox.Show("No se ha capturado el monto del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                txtimporte.Focus();
                                return;
                            }
                        }
                        if (CBtipoflete.SelectedIndex == -1)
                        {
                            MessageBox.Show("No se ha capturado el tipo del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            CBtipoflete.Focus();
                            return;
                        }
                        if (CBmonfle.SelectedIndex == -1)
                        {
                            MessageBox.Show("No se ha capturado la moneda del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            CBmonfle.Focus();
                            return;
                        }
                    }

                    if (Convert.ToDecimal(txtimporte.Text) > 0)
                    {
                        if (txttalonemb.Text.Trim() == "")
                        {
                            MessageBox.Show("No se ha capturado el talon de embarque", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            txttalonemb.Focus();
                            return;
                        }
                    }


                    if (txttalonemb.Text.Trim() != "")
                    {
                        if (txtcompleped.Text.Trim() == "")
                        {
                            if (Convert.ToDecimal(txtimporte.Text) == 0)
                            {
                                MessageBox.Show("No se ha capturado el monto del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                txtimporte.Focus();
                                return;
                            }
                        }
                        if (CBtipoflete.SelectedIndex == -1)
                        {
                            MessageBox.Show("No se ha capturado el tipo del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            CBtipoflete.Focus();
                            return;
                        }
                        if (CBmonfle.SelectedIndex == -1)
                        {
                            MessageBox.Show("No se ha capturado la moneda del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            CBmonfle.Focus();
                            return;
                        }
                    }

                    if (Convert.ToDecimal(txtimporte.Text) > 0)
                    {
                        if (txttalonemb.Text.Trim() == "")
                        {
                            MessageBox.Show("No se ha capturado el talon de embarque", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            txttalonemb.Focus();
                            return;
                        }
                    }

                    if (CBTransportista.SelectedIndex == -1)
                    {
                        MessageBox.Show("Favor de elegir transportista", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (DGV2.Rows.Count == 0)
                    {
                        MessageBox.Show("La orden de venta no tiene productos", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        //if (MessageBox.Show("¿Está seguro de guardar?", "SIPGAB", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                        //{
                        for (int i = 0; i < DGV2.Rows.Count; i++)
                        {
                            if (Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) == 0)
                            {
                                MessageBox.Show("El producto: " + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + " no tiene precio\n\r Favor de agregarlo", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            if (Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) == 0)
                            {
                                MessageBox.Show("El producto: " + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + " no tiene número de cajas\n\r Favor de agregarlo", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        //se muestra en pantalla la orden de venta
                        imprimir();
                        //imprimir(Utilerias.Nombre_impresora);

                        if (MessageBox.Show("Desea realizar algún cambio en la orden de venta", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                        {
                            return;
                        }
                        else
                        {
                            //MessageBox.Show("Inicia proceso");
                            //Stopwatch st = new Stopwatch();
                            //st.Start();
                            validar();
                            thisConnection.Open();

                            if (moneda == "PESOS")
                                txtmn.Text = "0";

                            if (moneda == "USD")
                                txtusd.Text = "0";

                            if (moneda == "EUROS")
                                txteuro.Text = "0";

                            for (int a = 0; a < DGV2.Rows.Count; a++)
                            {
                                if (moneda == "PESOS")
                                    txtmn.Text = (Convert.ToDecimal(txtmn.Text) + Convert.ToDecimal(DGV2.Rows[a].Cells["total"].Value)).ToString();

                                if (moneda == "USD")
                                    txtusd.Text = (Convert.ToDecimal(txtusd.Text) + Convert.ToDecimal(DGV2.Rows[a].Cells["total"].Value)).ToString();

                                if (moneda == "EUROS")
                                    txteuro.Text = (Convert.ToDecimal(txteuro.Text) + Convert.ToDecimal(DGV2.Rows[a].Cells["total"].Value)).ToString();
                            }


                            //trae el vendedor 
                            cmnd2 = thisConnection.CreateCommand();
                            cmnd2.CommandText = "SELECT tcv.vend_clave, tcv.vend_nombre FROM tb_cat_cliente tcc, tb_cat_vendedores tcv WHERE tcc.cnte_clave = '" + txtcliente.Text + "' AND tcv.vend_nombre = tcc.vend_clave";
                            reader2 = cmnd2.ExecuteReader();
                            while (reader2.Read())
                            {
                                vend_clave = reader2.GetSqlString(0).ToString();
                                vend_nombre = reader2.GetSqlString(1).ToString();
                            }
                            reader2.Close();
                            if (vend_clave.Trim() == "")
                                vend_clave = Utilerias.Class1.Usu_login;

                            try
                            {
                                btnGuardar.Enabled = false;
                                if (txtmn.Text != string.Empty)
                                    resulta = Convert.ToDecimal(txtmn.Text);

                                if (txtusd.Text != string.Empty)
                                    resulta = Convert.ToDecimal(txtusd.Text);

                                if (txteuro.Text != string.Empty)
                                    resulta = Convert.ToDecimal(txteuro.Text);

                                if (cbIVA.Checked == true)
                                    iva_flete = "S";

                                if (cbpagoGAB.Checked == true)
                                    flete_gab = "S";

                                if (txtcompleped.Text.Trim() == "")
                                    pedorigen = 0;
                                else
                                    pedorigen = Convert.ToInt32(txtcompleped.Text);

                                if (CBmoneda.SelectedIndex != -1)
                                    monedaflete = CBmoneda.SelectedItem.ToString().Trim();

                                if (CBtipoflete.SelectedIndex != -1)
                                    tipoflete = CBtipoflete.SelectedItem.ToString().Trim();
                                string mSubCli = "";
                                foreach (DataRow dr in subcnte.Select("entr_nombre = '" + cbdestino.Text + "'"))
                                {
                                    mSubCli = dr["entr_clave"].ToString().Trim();
                                }

                                //Inicia el Calcular la suma de split antes de imprimir
                                decimal resultacal = 0;
                                decimal enteroscal = 0;
                                decimal decimalescal = 0;
                                decimal sumasplitcal = 0;
                                for (int i = 0; i < DGV2.Rows.Count; i++)
                                {
                                    decimal prod_num_tarimascal = 0;
                                    //DGV2.Rows[i].Cells["clave_producto"].Value
                                    cmnd2 = thisConnection.CreateCommand();
                                    cmnd2.CommandText = "select prod_num_tarimas from tb_cat_producto where prod_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "'";
                                    reader2 = cmnd2.ExecuteReader();
                                    while (reader2.Read())
                                    {
                                        prod_num_tarimascal = reader2.GetSqlDecimal(0).Value;
                                    }
                                    reader2.Dispose();
                                    if (prod_num_tarimascal == 0)
                                    {
                                        resultacal = 0;
                                    }
                                    else
                                    {
                                        resultacal = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) / prod_num_tarimascal;
                                    }

                                    enteroscal = Math.Truncate(resultacal);
                                    decimalescal = enteroscal - resultacal;
                                    sumasplitcal = sumasplitcal + decimalescal;
                                }

                                sumasplitcal = Convert.ToDecimal((sumasplitcal * -1).ToString("0.0"));

                                //Termina el calculo de la suma de split
                                destino = "";
                                cmnd2 = thisConnection.CreateCommand();
                                //cmnd2.CommandText = "insert into tb_mstr_pedidos_exp (pdn_fecha, cnte_clave, vendedor, pdn_estatus, pdn_tipo, pdn_monto_total, pdn_observacion, prov_clave, pdn_transporte, pdn_pedsigma, pdn_hora, pdn_elaboro, placacaja, pdn_sello, um_clave, pdn_monto_transporte) values ('" + Convert.ToDateTime(txtfecha.Text).ToString("s") + "','" + txtcliente.Text + "','" + vend_nombre + "', ' ', 'EXP', " + resulta + ",'" + txtobs.Text + "','" + txtclavetrans.Text + "','" + txttalonemb.Text + "','" + txtpedsigma.Text + "','" + DateTime.Now.ToString("HH:mm:ss") + "','" + vend_clave + "','" + txtcajatrailer.Text + "','" + txtsello.Text + "','" + moneda + "','" + txtimporte.Text + "') SELECT SCOPE_IDENTITY()";
                                cmnd2.CommandText = "insert into tb_mstr_pedidos_exp (pdn_fecha, cnte_clave, vendedor, pdn_estatus, pdn_tipo, pdn_monto_total, pdn_observacion, prov_clave, " +
                                                    "pdn_transporte, pdn_pedsigma, pdn_hora, pdn_elaboro, placacaja, pdn_sello, um_clave, pdn_monto_transporte, cve_subcli, " +
                                                    "pdn_folio_web, ryandigital, nc_folio, pdn_situacion, pdn_entcve1, pdn_horent1, pdn_entcve2, pdn_horent2, pdn_entcve3, pdn_horent3, " +
                                                    "pdn_autmerma, pdn_surtido, ryan, pdn_iva, pdn_tipoflete, pdn_monedaflete, pdn_pagofletegab, pdn_ivaflete, pdn_pedorigen, pdn_destino, pdn_tractor, " +
                                                    "pdn_tractorplaca, pdn_cajatrailer, pdn_transoperador, pdn_cajacapacidad, pdn_sumsplit) values " +
                                                    "('" + Convert.ToDateTime(txtfecha.Text).ToString("s") + "','" + txtcliente.Text + "','" + vend_nombre + "', 'T', 'EXP', " +
                                                    "" + Convert.ToDecimal(resulta) + ",'" + txtobs.Text + "','" + txtclavetrans.Text + "','" + txttalonemb.Text + "','" + txtpedsigma.Text.Trim() + "', " +
                                                    "'" + DateTime.Now.ToString("HH:mm:ss") + "','" + Utilerias.Class1.Usu_login + "','" + txtcajatrailer.Text + "','" + txtsello.Text + "', " +
                                                    "'" + moneda + "', " + Convert.ToDecimal(txtimporte.Text) + ", '" + mSubCli + "', '', '', '', '', '', '" + DtHrCita.Text.ToString() + "', '', '', '', '', '', '', '', 0, " +
                                                    "'" + tipoflete + "', '" + monedaflete + "', '" + flete_gab + "', '" + iva_flete + "', " + pedorigen + ", '" + destino.Trim() + "', " +
                                                    "'" + TxtTractor.Text + "','" + TxtCjTracto.Text + "','" + TxtCjTra.Text + "','" + TxtChofer.Text + "','" + TxtCapaCaja.Text + "', '" + sumasplitcal + "') " +
                                                    "SELECT SCOPE_IDENTITY()";
                                /*cmnd2.CommandText = "insert into tb_mstr_pedidos_exp (pdn_fecha, cnte_clave, vendedor, cve_subcli, pdn_estatus, pdn_tipo, pdn_monto_total, pdn_folio_web, "+
                                                    "pdn_observacion, prov_clave, pdn_transporte, pdn_pedsigma, pdn_hora, pdn_elaboro, placacaja, pdn_sello, ryandigital, pdn_iva, um_clave, "+
                                                    "pdn_monto_transporte, nc_folio, pdn_situacion, pdn_entcve1, pdn_horent1, pdn_entcve2, pdn_horent2, pdn_entcve3, pdn_horent3, "+
                                                    "pdn_autmerma, pdn_surtido, ryan) values ('" + txtfecha.Text + "','" + txtcliente.Text.Trim() + "', '" + vend_nombre.Trim() + "', "+
                                                    "'', ' ','EXP', " + resulta + ",'', '" + txtobs.Text.Trim() + "','" + txtclavetrans.Text.Trim() + "', "+
                                                    "'" + txttalonemb.Text.Trim() + "', '" + txtpedsigma.Text.Trim() + "','" + DateTime.Now.ToString("HH:mm:ss") + "', "+
                                                    "'" + Utilerias.Usu_login.Trim() + "','" + txtcajatrailer.Text.Trim() + "','" + txtsello.Text.Trim() + "', "+
                                                    "'" + ryandigital.Trim() + "', 0 , '" + moneda + "','" + txtimporte.Text.Trim() + "', '', '', '', '', '', '', '', '', "+
                                                    "'" + DateTime.Now.ToString() + "','', '') SELECT SCOPE_IDENTITY()";*/
                                ultimofolio = Convert.ToString(cmnd2.ExecuteScalar()).Trim();

                                //verifica el ultimo folio que se guardo
                                //cmnd2 = new SqlCommand("SELECT IDENT_CURRENT('tb_mstr_pedidos_exp')", thisConnection);
                                //reader2 = cmnd2.ExecuteReader();
                                //while (reader2.Read())
                                //{
                                //    ultimofolio = reader2.GetSqlValue(0).ToString();
                                //}
                                //reader2.Close();

                                //verifica si el folio cambio y se lo muestra al usuario si cambio
                                if (ultimofolio.ToString().Trim() != txtfolio.Text.Trim())
                                {
                                    MessageBox.Show("El folio de esta orden de venta a cambiado a " + ultimofolio, "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    txtfolio.Text = ultimofolio;
                                }
                                string Cad = cmnd2.CommandText.ToString();
                                Cad = Cad.Replace("'", "");
                                Utilerias.Class1.registrar_movimiento3(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1.1", ultimofolio, "Alta MSTR EXP P-" + ultimofolio + " Imp $" + txtusd.Text + " Moneda: " + CBmonfle.Text, "SIPGAB", ultimofolio, "Tot Gen $" + txtusd.Text + " VS Tot Det $ " + txtimporte.Text + " " + Cad);
                                //MyConnection.Open();

                                //OleDbCommand dbCmdNull1 = MyConnection.CreateCommand();
                                //dbCmdNull1.CommandText = "SET NULL OFF";
                                //dbCmdNull1.ExecuteNonQuery();

                                //cmd1.CommandText = "INSERT INTO tb_mstr_pedidos_nal ([pdn_folio], [cnte_clave],  [pdn_fecha], [vendedor], [pdn_monto_total], [pdn_estatus], [pdn_observacion], [pdn_tipo], [pdn_transporte], [prov_clave], [pdn_monto_transporte], [pdn_pedsigma], [pdn_autmerma], [pdn_hora], [pdn_elaboro], [pdn_sello], [um_clave], [placacaja]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
                                //cmd1.Parameters.Add("@param0", OleDbType.Char).Value = ultimofolio.Trim();                                    
                                //cmd1.Parameters.Add("@param1", OleDbType.Char).Value = txtcliente.Text.Trim();                                    
                                //cmd1.Parameters.Add("@param2", OleDbType.DBDate).Value = txtfecha.Text.Trim();
                                //cmd1.Parameters.Add("@param3", OleDbType.Char).Value = vend_nombre.Trim();
                                //cmd1.Parameters.Add("@param4", OleDbType.Numeric).Value = resulta;
                                //cmd1.Parameters.Add("@param5", OleDbType.Char).Value = " ";
                                //cmd1.Parameters.Add("@param6", OleDbType.Char).Value = txtobs.Text.Trim();
                                //cmd1.Parameters.Add("@param7", OleDbType.Char).Value = "EXP";
                                //cmd1.Parameters.Add("@param8", OleDbType.Char).Value = txttalonemb.Text.Trim();
                                //cmd1.Parameters.Add("@param9", OleDbType.Char).Value = txtclavetrans.Text.Trim();
                                //cmd1.Parameters.Add("@param10", OleDbType.Numeric).Value = txtimporte.Text.Trim();
                                //cmd1.Parameters.Add("@param11", OleDbType.Char).Value = txtpedsigma.Text.Trim();
                                //cmd1.Parameters.Add("@param12", OleDbType.Char).Value = DateTime.Now.ToString();
                                //cmd1.Parameters.Add("@param13", OleDbType.Char).Value = DateTime.Now.ToString("HH:mm:ss");
                                //cmd1.Parameters.Add("@param14", OleDbType.Char).Value = Utilerias.Usu_login;
                                //cmd1.Parameters.Add("@param15", OleDbType.Char).Value = txtsello.Text.Trim();
                                //cmd1.Parameters.Add("@param16", OleDbType.Char).Value = moneda.Trim();                                    
                                //cmd1.Parameters.Add("@param17", OleDbType.Char).Value = txtcajatrailer.Text.Trim();


                                //cmd1.Connection = MyConnection;
                                //cmd1.ExecuteNonQuery();
                                //MyConnection.Close();


                                //inserta los productos en la tb_det_pedidos
                                decimal tot = 0;
                                string mProd = "";
                                for (int i = 0; i < DGV2.Rows.Count; i++)
                                {
                                    if ((Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) > 0))
                                    {
                                        //si el tipo de moneda es USD o EUROS se pone en 0 los MN
                                        if (CBmoneda.SelectedIndex == 1 || CBmoneda.SelectedIndex == 2)
                                        {
                                            moneyusd = decimal.Parse(DGV2.Rows[i].Cells["precios"].Value.ToString());
                                            moneymn = 0;
                                        }
                                        //si el tipo de moneda es MN se pone en 0 USD
                                        if (CBmoneda.SelectedIndex == 0)
                                        {
                                            moneymn = decimal.Parse(DGV2.Rows[i].Cells["precios"].Value.ToString());
                                            moneyusd = 0;
                                        }

                                        string tran = Convert.ToString(DGV2.Rows[i].Cells["transito"].Value).Trim();
                                        string trans = "N";
                                        if (tran == "True")
                                            trans = "S";

                                        cmnd2 = thisConnection.CreateCommand();
                                        cmnd2.CommandText = "insert into tb_det_pedidos (pdn_folio, nrenglon, lin_clave, prod_clave, pdn_num_unidades, pdn_pedimento, pdn_fechap, " +
                                                            "pdn_aduana, pdn_precio_mn, pdn_tipo, " +
                                                            "pdn_precio_usd, pdn_estatus, pdn_lote, pdn_cantxcaja, pdn_pre_cad, pdn_observaciones, pdn_tipotar, pdn_ordventa, " +
                                                            "pdn_transito, pdn_subcli, pdn_modif, pdn_descrip_especial, pdn_num_ped, pdn_fecha_cad) values ('" + ultimofolio + "','" + (i + 1).ToString() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["linea"].Value).Trim() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value).Trim() + "', " +
                                                            "'" + Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) + "', '', '', '', '" + moneymn + "','EXP', " +
                                                            "'" + moneyusd + "', 'T', '', '', '', '" + Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value).Trim() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["tipo_tar"].Value).Trim() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["po_number"].Value) + "', '" + trans + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["subcnte_clave"].Value).Trim() + "', '', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["vta_especial"].Value) + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["pedimento"].Value) + "', " + FechaCadSql(DGV2.Rows[i]) + ")";
                                        reader2 = cmnd2.ExecuteReader();
                                        reader2.Close();
                                        thisConnection.Close();
                                        thisConnection.Open();
                                        tot += Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) * Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value);
                                        mProd += DGV2.Rows[i].Cells["clave_producto"].Value.ToString() + " " + DGV2.Rows[i].Cells["Nombre"].Value.ToString() + " Cant: " + DGV2.Rows[i].Cells["cajas"].Value.ToString() + " Pre $ " + DGV2.Rows[i].Cells["precios"].Value.ToString() + System.Environment.NewLine;
                                        Utilerias.Class1.registrar_movimiento3(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1.1", ultimofolio, "Alta DET EXP P-" + ultimofolio + " Imp $" + txtusd.Text + " Moneda: " + CBmonfle.Text, "SIPGAB", ultimofolio, "Tot Gen $" + txtusd.Text + " VS Tot Det $ " + tot + " " + mProd);
                                        //MyConnection.Open();

                                        //dbCmdNull1 = MyConnection.CreateCommand();
                                        //dbCmdNull1.CommandText = "SET NULL OFF";
                                        //dbCmdNull1.ExecuteNonQuery();
                                        //cmd1.Dispose();

                                        //if (Convert.ToString(DGV2.Rows[i].Cells["fecha_ped"].Value).Trim() == "")
                                        //{
                                        //    cmd1.CommandText = "insert into tb_det_pedidos_temporal ([pdn_folio], [nrenglon], [lin_clave], [prod_clave], [pdn_num_unidades], [pdn_pedimento], [pdn_aduana], [pdn_precio_mn], [pdn_tipo], [pdn_precio_usd]) values (?,?,?,?,?,?,?,?,?,?)";
                                        //}
                                        //else
                                        //{
                                        //    cmd1.CommandText = "insert into tb_det_pedidos_temporal ([pdn_folio], [nrenglon], [lin_clave], [prod_clave], [pdn_num_unidades], [pdn_pedimento], [pdn_fechap], [pdn_aduana], [pdn_precio_mn], [pdn_tipo], [pdn_precio_usd]) values (?,?,?,?,?,?,?,?,?,?,?)";
                                        //    cmd1.Parameters.Add("@Nombre6", OleDbType.Date).Value = DGV2.Rows[i].Cells["fecha_ped"].Value.ToString();
                                        //}

                                        //cmd1.Parameters.Add("@Nombre0", OleDbType.Char).Value = ultimofolio;
                                        //cmd1.Parameters.Add("@Nombre1", OleDbType.Char).Value = (1 + i).ToString();
                                        //cmd1.Parameters.Add("@Nombre2", OleDbType.Char).Value = DGV2.Rows[i].Cells["linea"].Value.ToString();
                                        //cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = DGV2.Rows[i].Cells["clave_producto"].Value.ToString();
                                        //cmd1.Parameters.Add("@Nombre4", OleDbType.Numeric, 10, "pdn_num_unidades");
                                        //cmd1.Parameters["@Nombre4"].Value = float.Parse(DGV2.Rows[i].Cells["cajas"].Value.ToString());
                                        //cmd1.Parameters.Add("@Nombre5", OleDbType.Char).Value = DGV2.Rows[i].Cells["pedimento"].Value.ToString();
                                        //cmd1.Parameters.Add("@Nombre7", OleDbType.Char).Value = DGV2.Rows[i].Cells["aduana"].Value.ToString();
                                        //cmd1.Parameters.Add("@Nombre8", OleDbType.Numeric).Value = moneymn;
                                        //cmd1.Parameters.Add("@Nombre9", OleDbType.Char).Value = "EXP";
                                        //cmd1.Parameters.Add("@Nombre10", OleDbType.Numeric).Value = moneyusd;
                                        ////cmd1.Parameters.Add("@Nombre13", OleDbType.Char).Value = DGV2.Rows[i].Cells["pza_kg"].Value.ToString();
                                        ////cmd1.Parameters.Add("@Nombre14", OleDbType.Char).Value = DGV2.Rows[i].Cells["precio_uni"].Value.ToString();

                                        //cmd1.Connection = MyConnection;
                                        //cmd1.ExecuteNonQuery();
                                        //MyConnection.Close();
                                        //cmd1.Dispose();
                                    }
                                }//for
                                //st.Stop();
                                //MessageBox.Show("Termina proceso, con un tiempo de: " + st.ElapsedMilliseconds);
                                //st.Reset();

                                if (MessageBox.Show("¿Desea mandar a inventario la orden de venta?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                {
                                    cmnd2 = thisConnection.CreateCommand();
                                    cmnd2.CommandText = "update tb_mstr_pedidos_exp set pdn_estatus = 'T' where pdn_folio = '" + txtfolio.Text + "'";
                                    reader2 = cmnd2.ExecuteReader();
                                    reader2.Close();

                                    //MyConnection.Open();
                                    //cmd1.CommandType = CommandType.Text;
                                    //cmd1.CommandText = "update tb_mstr_pedidos_nal set pdn_estatus = 'T' where pdn_folio = '" + ultimofolio + "' and pdn_tipo = 'EXP'";
                                    //cmd1.Connection = MyConnection;
                                    //cmd1.ExecuteNonQuery();
                                    //MyConnection.Close();

                                    cmnd2 = thisConnection.CreateCommand();
                                    cmnd2.CommandText = "update tb_det_pedidos set pdn_estatus = 'T' where pdn_folio= '" + ultimofolio + "' and pdn_tipo = 'EXP'";
                                    reader2 = cmnd2.ExecuteReader();
                                    reader2.Close();

                                    //MyConnection.Open();
                                    //cmd1.CommandType = CommandType.Text;
                                    //cmd1.CommandText = "update tb_det_pedidos_temporal set pdn_estatus = 'T' where pdn_folio='" + ultimofolio + "' and pdn_tipo = 'EXP'";
                                    //cmd1.Connection = MyConnection;
                                    //cmd1.ExecuteNonQuery();
                                    //MyConnection.Close();

                                    for (int i = 0; i < DGV2.Rows.Count; i++)
                                    {
                                        if (Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) > 0 && (Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) > 0))
                                        {
                                            Fn_peso_unitario(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value));
                                            var_peso_unitario = var_dec_peso_total;

                                            Fn_invpt_salidas(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value), Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value), var_peso_unitario);
                                        }
                                    }//for
                                    thisConnection.Close();
                                    if (MessageBox.Show("¿Desea imprimir esta orden de venta?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                    {
                                        imprimir(Utilerias.Class1.Nombre_impresora);
                                    }
                                }
                                else
                                {
                                    thisConnection.Close();
                                }
                                //Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1.1", ultimofolio, "NUEVA ORDEN DE VENTA DE EXP CON EL FOLIO " + ultimofolio, "SIPGAB");
                                // se cambio la funcion de guardar para obtener mas informacion en el reg de mov  19 jun 2020 RCC
                                Utilerias.Class1.registrar_movimiento3(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1.1", ultimofolio, "O. VTA EXP F-" + ultimofolio + " Imp $" + txtusd.Text + " Moneda: " + CBmonfle.Text, "SIPGAB", ultimofolio, "Tot Gen $" + txtusd.Text + " VS Tot Det $ " + tot + " " + mProd);
                            }
                            catch (SqlException ex)
                            {
                                thisConnection.Close();
                                Utilerias.Class1.SendMail("ricardo.cortes@mrlucky.com.mx", "ricardo.cortes", "rcedillo", ex.ToString());
                                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                            catch (Exception ex1)
                            {
                                thisConnection.Close();
                                Utilerias.Class1.SendMail("ricardo.cortes@mrlucky.com.mx", "ricardo.cortes", "rcedillo", ex1.ToString());
                                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                        }
                        //}
                    }
                    break;
                #endregion

                #region baja
                case 2:
                    if (MessageBox.Show("¿Esta seguro de cancelar la orden de venta?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        try
                        {
                            thisConnection.Open();
                            //cambia a C el estatus y surtido a S de la orden 
                            cmnd2 = thisConnection.CreateCommand();
                            cmnd2.CommandText = "update tb_mstr_pedidos_exp set pdn_estatus = 'C', placacaja = '', pdn_pedorigen = '0', pdn_surtido='S' where pdn_folio = '" + txtfolio.Text + "'";
                            reader2 = cmnd2.ExecuteReader();
                            reader2.Close();

                            //MyConnection.Open();
                            //cmd1.CommandType = CommandType.Text;
                            //cmd1.CommandText = "update tb_mstr_pedidos_nal set pdn_estatus = 'C', pdn_surtido='S' where pdn_folio = '" + txtfolio.Text + "' and pdn_tipo = 'EXP'";
                            //cmd1.Connection = MyConnection;
                            //cmd1.ExecuteNonQuery();
                            //MyConnection.Close();

                            //modifica el inventario pt
                            cmnd2 = thisConnection.CreateCommand();
                            cmnd2.CommandText = "select lin_clave, prod_clave, pdn_num_unidades, pdn_precio_mn, pdn_precio_usd from tb_det_pedidos where pdn_folio = '" + txtfolio.Text + "' and pdn_tipo = 'EXP'";
                            reader2 = cmnd2.ExecuteReader();
                            while (reader2.Read())
                            {
                                lin_clave = reader2.GetSqlString(0).ToString();
                                prod_clave = reader2.GetSqlString(1).ToString();
                                pdn_num_unidades = reader2.GetSqlDecimal(2).Value;
                                pdn_precio_mn = reader2.GetSqlDecimal(3).Value;
                                pdn_precio_usd = reader2.GetSqlDecimal(4).Value;

                                Fn_peso_unitario(lin_clave, prod_clave);
                                var_peso_unitario = var_dec_peso_total;

                                //Fn_invpt_salidas(lin_clave, prod_clave, pdn_num_unidades, var_peso_unitario);
                                Fn_invpt_entradas(lin_clave, prod_clave, pdn_num_unidades, var_peso_unitario);
                            }
                            reader2.Close();


                            cmnd3 = thisConnection.CreateCommand();
                            cmnd3.CommandText = "update tb_det_pedidos set pdn_estatus= 'C' where pdn_folio='" + txtfolio.Text + "' and pdn_tipo = 'EXP'";
                            reader3 = cmnd3.ExecuteReader();
                            reader3.Close();

                            //MyConnection.Open();
                            //cmd1.CommandType = CommandType.Text;
                            //cmd1.CommandText = "update tb_det_pedidos_temporal set pdn_estatus= 'C' where pdn_folio='" + txtfolio.Text + "' and pdn_tipo = 'EXP'";
                            //cmd1.Connection = MyConnection;
                            //cmd1.ExecuteNonQuery();
                            //MyConnection.Close();                                                         
                            thisConnection.Close();
                            Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "B", "4.1.1", txtfolio.Text, "CANCELACION ORDEN DE VENTA DE EXP CON EL FOLIO " + txtfolio.Text, "SIPGAB");
                        }
                        catch (SqlException ex)
                        {
                            thisConnection.Close();
                            Utilerias.Class1.SendMail("ricardo.cortes@mrlucky.com.mx", "ricardo.cortes", "rcedillo", ex.ToString());
                            MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        catch (Exception ex1)
                        {
                            thisConnection.Close();
                            Utilerias.Class1.SendMail("ricardo.cortes@mrlucky.com.mx", "ricardo.cortes", "rcedillo", ex1.ToString());
                            MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                    }
                    break;
                #endregion

                #region actualizar
                case 3:

                    if (ValidarCambiosEmbarqueOrden() != "NO")
                    {
                        string estadoorden = ValidarOrdenVentaEmbarque();
                        if (estadoorden == "Iniciado")
                        {
                            Utilerias.Class1.registrar_movimiento3(DateTime.Now,
                                        Environment.MachineName,
                                        Utilerias.Class1.Usu_login,
                                        "C",
                                        "4.1",
                                        txtfolio.Text,
                                        "La Orden de venta EXP tiene embarque Iniciado, Considere avisar a Embarques de Cualquier cambio y considere que los tiempos de Carga seran afectados",
                                        "SIPGAB",
                                        txtfolio.Text,
                                        "Se muestra el aviso al usuario");
                            MessageBox.Show("La Orden de venta tiene embarque Iniciado, Considere avisar a Embarques de Cualquier cambio y considere que los tiempos de Carga seran afectados", "Carga Iniciada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else if (estadoorden == "Finalizado")
                        {
                            MessageBox.Show("La Orden de venta tiene embarque Terminado y cerrado. No se pueden realizar Modificaciones a los productos hasta que la orden sea abierta", "Carga Terminada", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                        }
                    }

                    if (DtHrCita.Text.ToString().Trim() == "00:00")
                    {
                        MessageBox.Show("La hora no puede ser 00:00, Favor de poner un horario de salida real para considerar el armado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DtHrCita.Focus();
                        return;
                    }

                    if (cbdestino.SelectedIndex != -1)
                    {
                        if (Convert.ToString(cbdestino.SelectedItem).Trim() == "MCALLEN, TX.")
                        {
                            for (int i = 0; i < DGV2.Rows.Count; i++)
                            {
                                if ((Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("APIO")) || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("LECHUGA"))
                                    || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("KALE")) || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("SWISS CHARD"))
                                    || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("RED SWISS CHARD")) || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("CILANTRO"))
                                    || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("RED KALE")))
                                    leyenda1 = "ASEGURARSE QUE EN ESTE EMBARQUE HAYA UNA DISTANCIA DE 18 PULGADAS ENTRE LA PARTE ALTA DEL PALLET Y EL TECHO PARA LOS QUE CRUZAN POR " + cbdestino.SelectedItem.ToString().Trim();
                            }
                        }
                    }
                    if (cbdestino.SelectedIndex == -1)
                        destino = "";
                    else
                        destino = cbdestino.SelectedItem.ToString().Trim();
                    if (cbpagoGAB.Checked == true)
                    {
                        if (txttalonemb.Text.Trim() == "")
                        {
                            MessageBox.Show("No se ha capturado el talon de embarque", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            txttalonemb.Focus();
                            return;
                        }
                        if (txtcompleped.Text.Trim() == "")
                        {
                            if (Convert.ToDecimal(txtimporte.Text) == 0)
                            {
                                MessageBox.Show("No se ha capturado el monto del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                txtimporte.Focus();
                                return;
                            }
                        }
                    }

                    if (CBTransportista.SelectedIndex == -1)
                    {
                        MessageBox.Show("Favor de elegir transportista", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }


                    if (txttalonemb.Text.Trim() != "")
                    {
                        if (txtcompleped.Text.Trim() == "")
                        {
                            if (Convert.ToDecimal(txtimporte.Text) == 0)
                            {
                                MessageBox.Show("No se ha capturado el monto del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                txtimporte.Focus();
                                return;
                            }
                        }
                        if (CBtipoflete.SelectedIndex == -1)
                        {
                            MessageBox.Show("No se ha capturado el tipo del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            CBtipoflete.Focus();
                            return;
                        }
                        if (CBmonfle.SelectedIndex == -1)
                        {
                            MessageBox.Show("No se ha capturado la moneda del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            CBmonfle.Focus();
                            return;
                        }
                    }

                    if (Convert.ToDecimal(txtimporte.Text) > 0)
                    {
                        if (txttalonemb.Text.Trim() == "")
                        {
                            MessageBox.Show("No se ha capturado el talon de embarque", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            txttalonemb.Focus();
                            return;
                        }
                    }

                    if (DGV2.Rows.Count == 0)
                    {
                        MessageBox.Show("La orden de venta no tiene productos", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        //if (MessageBox.Show("¿Está seguro de guardar?", "SIPGAB", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                        //{

                        //if (cbpagoGAB.Checked == true)
                        //{
                        //    if (Convert.ToDecimal(txtimporte.Text) == 0)
                        //    {
                        //        MessageBox.Show("Favor de capturar el importe del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        //        txtimporte.Focus();
                        //        return;
                        //    }
                        //    if (txttalonemb.Text.Trim() == "")
                        //    {
                        //        MessageBox.Show("Favor de capturar el talon de embarque", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        //        txttalonemb.Focus();
                        //        return;
                        //    }
                        //}

                        for (int i = 0; i < DGV2.Rows.Count; i++)
                        {
                            if (Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) == 0)
                            {
                                MessageBox.Show("El producto: " + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + " no tiene precio\n\r Favor de agregarlo", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            if (Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) == 0)
                            {
                                MessageBox.Show("El producto: " + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + " no tiene número de cajas\n\r Favor de agregarlo", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        if (moneda == "PESOS")
                            txtmn.Text = "0";

                        if (moneda == "USD")
                            txtusd.Text = "0";

                        if (moneda == "EUROS")
                            txteuro.Text = "0";

                        for (int a = 0; a < DGV2.Rows.Count; a++)
                        {
                            if (moneda == "PESOS")
                                txtmn.Text = (Convert.ToDecimal(txtmn.Text) + Convert.ToDecimal(DGV2.Rows[a].Cells["total"].Value)).ToString();

                            if (moneda == "USD")
                                txtusd.Text = (Convert.ToDecimal(txtusd.Text) + Convert.ToDecimal(DGV2.Rows[a].Cells["total"].Value)).ToString();

                            if (moneda == "EUROS")
                                txteuro.Text = (Convert.ToDecimal(txteuro.Text) + Convert.ToDecimal(DGV2.Rows[a].Cells["total"].Value)).ToString();
                        }

                        //se muestra en pantalla la orden de venta
                        imprimir();

                        if (MessageBox.Show("Desea confirmar el cambio realizado", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                        {
                            validar();
                            try
                            {


                                //Inicia el Calcular la suma de split antes de imprimir
                                decimal resultacal = 0;
                                decimal enteroscal = 0;
                                decimal decimalescal = 0;
                                decimal sumasplitcal = 0;
                                for (int i = 0; i < DGV2.Rows.Count; i++)
                                {
                                    decimal prod_num_tarimascal = 0;
                                    //DGV2.Rows[i].Cells["clave_producto"].Value
                                    thisConnection.Open();
                                    cmnd2 = thisConnection.CreateCommand();
                                    cmnd2.CommandText = "select prod_num_tarimas from tb_cat_producto where prod_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "'";
                                    reader2 = cmnd2.ExecuteReader();
                                    while (reader2.Read())
                                    {
                                        prod_num_tarimascal = reader2.GetSqlDecimal(0).Value;
                                    }
                                    reader2.Dispose();
                                    if (prod_num_tarimascal == 0)
                                    {
                                        resultacal = 0;
                                    }
                                    else
                                    {
                                        resultacal = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) / prod_num_tarimascal;
                                    }
                                    thisConnection.Close();
                                    enteroscal = Math.Truncate(resultacal);
                                    decimalescal = enteroscal - resultacal;
                                    sumasplitcal = sumasplitcal + decimalescal;
                                }

                                sumasplitcal = Convert.ToDecimal((sumasplitcal * -1).ToString("0.0"));

                                //Termina el calculo de la suma de split

                                if (txtmn.Text != string.Empty)
                                    resulta = Convert.ToDecimal(txtmn.Text);

                                if (txtusd.Text != string.Empty)
                                    resulta = Convert.ToDecimal(txtusd.Text);

                                if (txteuro.Text != string.Empty)
                                    resulta = Convert.ToDecimal(txteuro.Text);

                                if (cbIVA.Checked == true)
                                    iva_flete = "S";
                                else
                                {
                                    iva_flete = "N";
                                    txtIVAfle.Text = "";
                                    txtretension.Text = "";
                                }

                                if (cbpagoGAB.Checked == true)
                                    flete_gab = "S";
                                else
                                    flete_gab = "N";

                                if (txtcompleped.Text.Trim() == "")
                                    pedorigen = 0;
                                else
                                    pedorigen = Convert.ToInt32(txtcompleped.Text);

                                if (CBmoneda.SelectedIndex != -1)
                                    monedaflete = CBmoneda.SelectedItem.ToString().Trim();

                                if (CBtipoflete.SelectedIndex != -1)
                                    tipoflete = CBtipoflete.SelectedItem.ToString().Trim();

                                thisConnection.Open();
                                //cmnd2 = thisConnection.CreateCommand();
                                //cmnd2.CommandText = "select * from tb_det_pedidos where pdn_folio = '" + txtfolio.Text + "' and pdn_tipo = 'EXP'";
                                //reader2 = cmnd2.ExecuteReader();
                                //while (reader2.Read())
                                //{
                                string mSubCli = "";
                                foreach (DataRow dr in subcnte.Select("entr_nombre = '" + cbdestino.Text + "'"))
                                {
                                    mSubCli = dr["entr_clave"].ToString().Trim();
                                }
                                if (destino.Trim().Length > 20)
                                    destino = destino.Substring(0, 19);
                                cmnd3 = thisConnection.CreateCommand();
                                cmnd3.CommandText = "update tb_mstr_pedidos_exp set pdn_fecha = '" + Convert.ToDateTime(txtfecha.Text).ToString("s") + "', " +
                                                    "pdn_monto_total = '" + resulta + "', pdn_observacion = '" + txtobs.Text + "', " +
                                                    "prov_clave = '" + txtclavetrans.Text + "', pdn_transporte = '" + txttalonemb.Text + "', " +
                                                    "pdn_pedsigma = '" + txtpedsigma.Text + "', pdn_hora='" + DateTime.Now.ToString("HH:mm:ss") + "', " +
                                                    "placacaja = '" + txtcajatrailer.Text + "', pdn_sello = '" + txtsello.Text + "', " +
                                                    "um_clave = '" + moneda + "', pdn_tipoflete = '" + tipoflete + "', pdn_monedaflete = '" + monedaflete + "', " +
                                                    "pdn_pagofletegab = '" + flete_gab + "', pdn_pedorigen = '" + pedorigen + "', pdn_ivaflete = '" + iva_flete + "', " +
                                                    "pdn_monto_transporte = '" + Convert.ToDecimal(txtimporte.Text) + "', cnte_clave = '" + txtcliente.Text + "', " +
                                                    "pdn_destino = '" + destino + "', pdn_horent1 = '" + DtHrCita.Text.ToString() + "', cve_subcli = '" + mSubCli + "', " +
                                                    "pdn_tractor = '" + TxtTractor.Text + "', pdn_tractorplaca = '" + TxtCjTracto.Text + "', pdn_cajatrailer = '" + TxtCjTra.Text + "', " +
                                                    "pdn_transOperador = '" + TxtChofer.Text + "', pdn_cajacapacidad = '" + TxtCapaCaja.Text + "', pdn_sumsplit = '" + sumasplitcal + "' " +
                                                    "where pdn_folio = '" + txtfolio.Text + "'";
                                reader3 = cmnd3.ExecuteReader();
                                reader3.Close();
                                thisConnection.Close();
                                Utilerias.Class1.registrar_movimiento3(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "M", "4.1.1", txtfolio.Text, "MODIFICACION A LA O. VTA EXP F-" + txtfolio.Text + " Imp $" + txtusd.Text + " Moneda: " + CBmonfle.Text, "SIPGAB", txtfolio.Text, cmnd3.CommandText.Replace("'", ""));
                                thisConnection.Open();
                                //MyConnection.Open();
                                //OleDbCommand dbCmdNull1 = MyConnection.CreateCommand();
                                //dbCmdNull1.CommandText = "SET NULL OFF";
                                //dbCmdNull1.ExecuteNonQuery();

                                //cmd1.CommandText = "update tb_mstr_pedidos_nal set pdn_fecha = ?, pdn_monto_total = ?, pdn_observacion = ?, prov_clave = ?, pdn_transporte = ?, pdn_pedsigma = ?, pdn_hora= ?, placacaja = ?, pdn_sello = ?, um_clave = ? where pdn_folio = '" + txtfolio.Text + "' and pdn_tipo ='EXP'";
                                //cmd1.Parameters.Add("@Nombre1", OleDbType.DBDate).Value = Convert.ToDateTime(txtfecha.Text).ToString("s");                                
                                //cmd1.Parameters.Add("@Nombre2", OleDbType.Numeric).Value = resulta;
                                //cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = txtobs.Text;
                                //cmd1.Parameters.Add("@Nombre4", OleDbType.Char).Value = txtclavetrans.Text;
                                //cmd1.Parameters.Add("@Nombre5", OleDbType.Char).Value = txttalonemb.Text;
                                //cmd1.Parameters.Add("@Nombre6", OleDbType.Char).Value = txtpedsigma.Text;
                                //cmd1.Parameters.Add("@Nombre7", OleDbType.Char).Value = Convert.ToString(DateTime.Now.ToString("HH:mm:ss"));
                                //cmd1.Parameters.Add("@Nombre8", OleDbType.Char).Value = txtcajatrailer.Text;
                                //cmd1.Parameters.Add("@Nombre9", OleDbType.Char).Value = txtsello.Text;                                                                
                                //cmd1.Parameters.Add("@Nombre10", OleDbType.Char).Value = moneda;

                                //cmd1.Connection = MyConnection;
                                //cmd1.ExecuteNonQuery();
                                //MyConnection.Close();
                                //cmd1.Dispose();
                                ////}
                                ////reader2.Close();

                                //MessageBox.Show("Inicia borrado");
                                //Stopwatch st = new Stopwatch();
                                //st.Start();
                                ////borra los productos del detalle de pedidos en fox
                                //MyConnection.Open();
                                //cmd1.CommandType = CommandType.Text;
                                //cmd1.CommandText = "delete from tb_det_pedidos_temporal where pdn_folio='" + txtfolio.Text + "'and pdn_tipo = 'EXP'";
                                //cmd1.Connection = MyConnection;
                                //cmd1.ExecuteNonQuery();
                                //MyConnection.Close();
                                //st.Stop();
                                //MessageBox.Show("Tiempo de borrado " + st.ElapsedMilliseconds);
                                //st.Reset();

                                //MessageBox.Show("Inicia for");
                                //st.Start();

                                cmnd2 = thisConnection.CreateCommand();
                                cmnd2.CommandText = "delete from tb_det_pedidos where pdn_folio = '" + txtfolio.Text + "' and pdn_tipo = 'EXP'";
                                reader2 = cmnd2.ExecuteReader();
                                reader2.Dispose();

                                #region recorre el DGV2 e inserta los productos en la tb_det_pedidos
                                decimal tot = 0;
                                string mProd = "";
                                for (int i = 0; i < DGV2.Rows.Count; i++)
                                {
                                    if (Convert.ToString(DGV2.Rows[i].Cells["vta_especial"].Value).Contains("'"))
                                        DGV2.Rows[i].Cells["vta_especial"].Value = Convert.ToString(DGV2.Rows[i].Cells["vta_especial"].Value).Replace("'", "");
                                    //consulta en la tb_det_pedidos si existe el producto
                                    //cmnd2 = thisConnection.CreateCommand();
                                    //cmnd2.CommandText = "select pdn_num_unidades, pdn_precio_mn, pdn_precio_usd, pdn_observaciones, pdn_tipotar, pdn_ordventa, pdn_subcli "+
                                    //                    "from tb_det_pedidos where pdn_folio = '" + txtfolio.Text + "' and "+
                                    //                    "lin_clave ='" + Convert.ToString(DGV2.Rows[i].Cells["linea"].Value) + "' and "+
                                    //                    "prod_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "' and pdn_tipo = 'EXP'";
                                    //reader2 = cmnd2.ExecuteReader();
                                    //while (reader2.Read())
                                    //{
                                    //    pdn_num_unidades = reader2.GetSqlDecimal(0).Value;
                                    //    pdn_precio_mn = reader2.GetSqlDecimal(1).Value;
                                    //    pdn_precio_usd = reader2.GetSqlDecimal(2).Value;
                                    //    pdn_pedimento = reader2.GetSqlString(3).ToString();//observaciones
                                    //    pdn_fechap = reader2.GetSqlString(4).ToString();//tipotar
                                    //    pdn_aduana = reader2.GetSqlString(5).ToString();//ordenventa
                                    //    pdn_cantxcaja = reader2.GetSqlString(6).ToString();//subcliente                                        

                                    //    cmnd3 = thisConnection.CreateCommand();

                                    //    //si el tipo de moneda es MN se pone en 0 los USD    
                                    if (CBmoneda.SelectedIndex == 0)
                                    {
                                        //        cmnd3.CommandText = "update tb_det_pedidos set pdn_lote=' ', pdn_num_unidades = '" + Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) + "', "+
                                        //                            "pdn_observaciones = '" + Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value) + "', "+
                                        //                            "pdn_tipotar='" + Convert.ToString(DGV2.Rows[i].Cells["tipo_tar"].Value) + "', "+
                                        //                            "pdn_ordventa ='" + Convert.ToString(DGV2.Rows[i].Cells["po_number"].Value) + "', "+
                                        //                            "pdn_precio_mn = '" + Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) + "', "+
                                        //                            "nrenglon = '" + (i + 1).ToString() + "', "+
                                        //                            "pdn_subcli= '" + Convert.ToString(DGV2.Rows[i].Cells["subcnte_clave"].Value) + "' "+
                                        //                            "pdn_estatus = 'T' " +
                                        //                            "where pdn_folio = '" + txtfolio.Text + "' and lin_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["linea"].Value) + "' "+
                                        //                            "and prod_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "' and pdn_tipo= 'EXP'";
                                        moneyusd = 0;
                                        moneymn = Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value);
                                    }
                                    //    //si el tipo de moneda es USD se pone en 0 MN
                                    if (CBmoneda.SelectedIndex == 1 || CBmoneda.SelectedIndex == 2)
                                    {
                                        //        cmnd3.CommandText = "update tb_det_pedidos set pdn_lote=' ', pdn_num_unidades = '" + Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) + "', " +
                                        //                            "pdn_observaciones = '" + Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value) + "', " +
                                        //                            "pdn_tipotar='" + Convert.ToString(DGV2.Rows[i].Cells["tipo_tar"].Value) + "', " +
                                        //                            "pdn_ordventa ='" + Convert.ToString(DGV2.Rows[i].Cells["po_number"].Value) + "', " +
                                        //                            "pdn_precio_usd = '" + Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) + "', " +
                                        //                            "nrenglon = '" + (i + 1).ToString() + "', " +
                                        //                            "pdn_subcli= '" + Convert.ToString(DGV2.Rows[i].Cells["subcnte_clave"].Value) + "' " +
                                        //                            "pdn_estatus = 'T' " +
                                        //                            "where pdn_folio = '" + txtfolio.Text + "' and lin_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["linea"].Value) + "' " +
                                        //                            "and prod_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "' and pdn_tipo= 'EXP'";
                                        moneymn = 0;
                                        moneyusd = Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value);
                                    }
                                    //    reader3 = cmnd3.ExecuteReader();
                                    //    reader3.Close();

                                    //MyConnection.Open();

                                    //dbCmdNull1 = MyConnection.CreateCommand();
                                    //dbCmdNull1.CommandText = "SET NULL OFF";
                                    //dbCmdNull1.ExecuteNonQuery();
                                    //cmd1.Dispose();

                                    //if (Convert.ToString(DGV2.Rows[i].Cells["fecha_ped"].Value).Trim() == "")
                                    //{
                                    //    //cmd1.CommandText = "insert into tb_det_pedidos_temporal ([pdn_folio], [nrenglon], [lin_clave], [prod_clave], [pdn_num_unidades], [pdn_pedimento], [pdn_aduana], [pdn_precio_mn], [pdn_tipo], [pdn_precio_usd], [pdn_estatus]) values (?,?,?,?,?,?,?,?,?,?,?)";
                                    //    cmd1.CommandText = "update tb_det_pedidos_temporal set pdn_num_unidades = ?, pdn_pedimento = ?, pdn_aduana = ?, nrenglon = ?,  pdn_precio_mn = ?, pdn_precio_usd = ? where pdn_folio = '" + txtfolio.Text + "' and lin_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["linea"].Value) + "' and prod_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "' and pdn_tipo= 'EXP'";
                                    //}
                                    //else
                                    //{
                                    //    //cmd1.CommandText = "insert into tb_det_pedidos_temporal ([pdn_folio], [nrenglon], [lin_clave], [prod_clave], [pdn_num_unidades], [pdn_pedimento], [pdn_fechap], [pdn_aduana], [pdn_precio_mn], [pdn_tipo], [pdn_precio_usd], [pdn_estatus]) values (?,?,?,?,?,?,?,?,?,?,?,?)";
                                    //    cmd1.CommandText = "update tb_det_pedidos_temporal set pdn_num_unidades = ?, pdn_pedimento = ?, pdn_fechap= ?, pdn_aduana = ?, nrenglon = ?,  pdn_precio_mn = ?, pdn_precio_usd = ? where pdn_folio = '" + txtfolio.Text + "' and lin_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["linea"].Value) + "' and prod_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "' and pdn_tipo= 'EXP'";
                                    //    cmd1.Parameters.Add("@Nombre6", OleDbType.Date).Value = DGV2.Rows[i].Cells["fecha_ped"].Value.ToString();
                                    //}

                                    ////cmd1.Parameters.Add("@Nombre0", OleDbType.Char).Value = txtfolio.Text;
                                    ////cmd1.Parameters.Add("@Nombre1", OleDbType.Char).Value = (1 + i).ToString();
                                    ////cmd1.Parameters.Add("@Nombre2", OleDbType.Char).Value = DGV2.Rows[i].Cells["linea"].Value.ToString();
                                    ////cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = DGV2.Rows[i].Cells["clave_producto"].Value.ToString();
                                    ////cmd1.Parameters.Add("@Nombre4", OleDbType.Numeric, 10, "pdn_num_unidades");
                                    ////cmd1.Parameters["@Nombre4"].Value = float.Parse(DGV2.Rows[i].Cells["cajas"].Value.ToString());
                                    ////cmd1.Parameters.Add("@Nombre5", OleDbType.Char).Value = DGV2.Rows[i].Cells["pedimento"].Value.ToString();
                                    ////cmd1.Parameters.Add("@Nombre7", OleDbType.Char).Value = DGV2.Rows[i].Cells["aduana"].Value.ToString();
                                    ////cmd1.Parameters.Add("@Nombre8", OleDbType.Numeric).Value = moneymn;
                                    ////cmd1.Parameters.Add("@Nombre9", OleDbType.Char).Value = "EXP";
                                    ////cmd1.Parameters.Add("@Nombre10", OleDbType.Numeric).Value = moneyusd;
                                    ////cmd1.Parameters.Add("@Nombre11", OleDbType.Char).Value = pdn_estatus.Trim();

                                    //cmd1.CommandType = CommandType.Text;
                                    //cmd1.Parameters.Add("@Nombre0", OleDbType.Numeric);
                                    //cmd1.Parameters["@Nombre0"].Value = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value);
                                    //cmd1.Parameters.Add("@Nombre1", OleDbType.Char);
                                    //cmd1.Parameters["@Nombre1"].Value = Convert.ToString(DGV2.Rows[i].Cells["pedimento"].Value).Trim();
                                    //cmd1.Parameters.Add("@Nombre3", OleDbType.Char);
                                    //cmd1.Parameters["@Nombre3"].Value = Convert.ToString(DGV2.Rows[i].Cells["aduana"].Value).Trim();                                                                            
                                    //string renglon = (i + 1).ToString();
                                    //cmd1.Parameters.Add("@Nombre7", OleDbType.Char, 2, "nrenglon");
                                    //cmd1.Parameters["@Nombre7"].Value = renglon;
                                    //cmd1.Parameters.Add("@Nombre6", OleDbType.Numeric).Value = moneymn;
                                    //cmd1.Parameters.Add("@Nombre8", OleDbType.Numeric).Value = moneyusd;

                                    //cmd1.Connection = MyConnection;
                                    //cmd1.ExecuteNonQuery();
                                    //MyConnection.Close();
                                    //cmd1.Dispose();

                                    //if (pdn_estatus.Trim() == "T")
                                    //{
                                    //    if (Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) < pdn_num_unidades)
                                    //    {
                                    //        pdn_num_unidades = pdn_num_unidades - Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value);
                                    //        Fn_peso_unitario(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value));
                                    //        var_peso_unitario = var_dec_peso_total;
                                    //        Fn_invpt_salidas(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value), pdn_num_unidades, var_peso_unitario);                                                
                                    //    }
                                    //    else
                                    //    {
                                    //        Fn_peso_unitario(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value));
                                    //        var_peso_unitario = var_dec_peso_total;
                                    //        Fn_invpt_entradas(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value), pdn_num_unidades, var_peso_unitario);
                                    //        Fn_invpt_salidas(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value), pdn_num_unidades, var_peso_unitario);
                                    //    }
                                    //}//pdn_estatus = T                                                                                                                                                                                                                              
                                    //}//while
                                    //if (reader2.HasRows == false)
                                    //{
                                    cmnd3 = thisConnection.CreateCommand();
                                    if (CBmoneda.SelectedIndex == 0)
                                        cmnd3.CommandText = "insert into tb_det_pedidos (pdn_folio, nrenglon, lin_clave, prod_clave, pdn_num_unidades, pdn_pedimento, pdn_fechap, " +
                                                            "pdn_aduana, pdn_precio_mn, pdn_tipo, " +
                                                            "pdn_precio_usd, pdn_estatus, pdn_lote, pdn_cantxcaja, pdn_pre_cad, pdn_observaciones, pdn_tipotar, pdn_ordventa, " +
                                                            "pdn_transito, pdn_subcli, pdn_modif, pdn_descrip_especial, pdn_num_ped, pdn_fecha_cad) values ('" + txtfolio.Text + "','" + (i + 1).ToString() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["linea"].Value).Trim() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value).Trim() + "', " +
                                                            "'" + Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) + "', '', '', '', '" + moneymn + "','EXP', " +
                                                            "'" + moneyusd + "', 'T', '', '', '', '" + Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value).Trim() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["tipo_tar"].Value).Trim() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["po_number"].Value) + "', '', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["subcnte_clave"].Value).Trim() + "', '', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["vta_especial"].Value) + "', '" + Convert.ToString(DGV2.Rows[i].Cells["pedimento"].Value) + "', " + FechaCadSql(DGV2.Rows[i]) + ")";

                                    if (CBmoneda.SelectedIndex == 1 || CBmoneda.SelectedIndex == 2)
                                        cmnd3.CommandText = "insert into tb_det_pedidos (pdn_folio, nrenglon, lin_clave, prod_clave, pdn_num_unidades, pdn_pedimento, pdn_fechap, " +
                                                            "pdn_aduana, pdn_precio_usd, pdn_tipo, " +
                                                            "pdn_precio_mn, pdn_estatus, pdn_lote, pdn_cantxcaja, pdn_pre_cad, pdn_observaciones, pdn_tipotar, pdn_ordventa, " +
                                                            "pdn_transito, pdn_subcli, pdn_modif, pdn_descrip_especial, pdn_num_ped, pdn_fecha_cad) values ('" + txtfolio.Text + "','" + (i + 1).ToString() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["linea"].Value).Trim() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value).Trim() + "', " +
                                                            "'" + Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) + "', '', '', '', '" + moneyusd + "','EXP', " +
                                                            "'" + moneymn + "', 'T', '', '', '', '" + Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value).Trim() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["tipo_tar"].Value).Trim() + "', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["po_number"].Value) + "', '', " +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["subcnte_clave"].Value).Trim() + "', ''," +
                                                            "'" + Convert.ToString(DGV2.Rows[i].Cells["vta_especial"].Value) + "', '" + Convert.ToString(DGV2.Rows[i].Cells["pedimento"].Value) + "', " + FechaCadSql(DGV2.Rows[i]) + ")";
                                    //cmnd3.CommandText = "insert into tb_det_pedidos (pdn_folio, nrenglon, lin_clave, prod_clave, pdn_num_unidades, pdn_precio_mn, pdn_precio_usd, pdn_tipo, pdn_pedimento, pdn_fechap, pdn_aduana, pdn_estatus, pdn_lote, pdn_observaciones) values ('" + txtfolio.Text + "','" + (1 + i).ToString() + "','" + Convert.ToString(DGV2.Rows[i].Cells["linea"].Value) + "','" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "','" + Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) + "', '0', '" + Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) + "','EXP', '" + Convert.ToString(DGV2.Rows[i].Cells["pedimento"].Value) + "','" + Convert.ToString(DGV2.Rows[i].Cells["fecha_ped"].Value) + "','" + Convert.ToString(DGV2.Rows[i].Cells["aduana"].Value) + "','" + pdn_estatus + "',' ', '" + Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value) + "')";
                                    //cmnd3.CommandText = "insert into tb_det_pedidos (pdn_folio, nrenglon, lin_clave, prod_clave, pdn_num_unidades, pdn_precio_mn, pdn_precio_usd, pdn_tipo, pdn_pedimento, pdn_fechap, pdn_aduana, pdn_cantxcaja, pdn_pre_cad) values ('" + txtfolio.Text + "','" + (1 + i).ToString() + "','" + Convert.ToString(DGV2.Rows[i].Cells["linea"].Value) + "','" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "','" + Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) + "', 0, '" + Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) + "', 'EXP', '" + Convert.ToString(DGV2.Rows[i].Cells["pedimento"].Value) + "','" + Convert.ToString(DGV2.Rows[i].Cells["fecha_ped"].Value) + "','" + Convert.ToString(DGV2.Rows[i].Cells["aduana"].Value) + "',' ',' ')";

                                    reader3 = cmnd3.ExecuteReader();
                                    reader3.Close();
                                    tot += Convert.ToDecimal(DGV2.Rows[i].Cells["precios"].Value) * Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value);
                                    mProd += DGV2.Rows[i].Cells["clave_producto"].Value.ToString() + " " + DGV2.Rows[i].Cells["Nombre"].Value.ToString() + " Cant: " + DGV2.Rows[i].Cells["cajas"].Value.ToString() + " Pre $ " + DGV2.Rows[i].Cells["precios"].Value.ToString() + System.Environment.NewLine;

                                    //MyConnection.Open();

                                    //dbCmdNull1 = MyConnection.CreateCommand();
                                    //dbCmdNull1.CommandText = "SET NULL OFF";
                                    //dbCmdNull1.ExecuteNonQuery();
                                    //cmd1.Dispose();
                                    //if (Convert.ToString(DGV2.Rows[i].Cells["fecha_ped"].Value).Trim() == "")
                                    //{
                                    //    cmd1.CommandText = "insert into tb_det_pedidos_temporal ([pdn_folio], [nrenglon], [lin_clave], [prod_clave], [pdn_num_unidades], [pdn_pedimento], [pdn_aduana], [pdn_precio_mn], [pdn_tipo], [pdn_precio_usd], [pdn_estatus]) values (?,?,?,?,?,?,?,?,?,?,?)";
                                    //}
                                    //else
                                    //{
                                    //    cmd1.CommandText = "insert into tb_det_pedidos_temporal ([pdn_folio], [nrenglon], [lin_clave], [prod_clave], [pdn_num_unidades], [pdn_pedimento], [pdn_fechap], [pdn_aduana], [pdn_precio_mn], [pdn_tipo], [pdn_precio_usd], [pdn_estatus]) values (?,?,?,?,?,?,?,?,?,?,?,?)";
                                    //    cmd1.Parameters.Add("@Nombre6", OleDbType.Date).Value = DGV2.Rows[i].Cells["fecha_ped"].Value.ToString();
                                    //}

                                    //cmd1.Parameters.Add("@Nombre0", OleDbType.Char).Value = txtfolio.Text;
                                    //cmd1.Parameters.Add("@Nombre1", OleDbType.Char).Value = (1 + i).ToString();
                                    //cmd1.Parameters.Add("@Nombre2", OleDbType.Char).Value = DGV2.Rows[i].Cells["linea"].Value.ToString();
                                    //cmd1.Parameters.Add("@Nombre3", OleDbType.Char).Value = DGV2.Rows[i].Cells["clave_producto"].Value.ToString();
                                    //cmd1.Parameters.Add("@Nombre4", OleDbType.Numeric, 10, "pdn_num_unidades");
                                    //cmd1.Parameters["@Nombre4"].Value = float.Parse(DGV2.Rows[i].Cells["cajas"].Value.ToString());
                                    //cmd1.Parameters.Add("@Nombre5", OleDbType.Char).Value = DGV2.Rows[i].Cells["pedimento"].Value.ToString();
                                    //cmd1.Parameters.Add("@Nombre7", OleDbType.Char).Value = DGV2.Rows[i].Cells["aduana"].Value.ToString();
                                    //cmd1.Parameters.Add("@Nombre8", OleDbType.Numeric).Value = moneymn;
                                    //cmd1.Parameters.Add("@Nombre9", OleDbType.Char).Value = "EXP";
                                    //cmd1.Parameters.Add("@Nombre10", OleDbType.Numeric).Value = moneyusd;
                                    //cmd1.Parameters.Add("@Nombre11", OleDbType.Char).Value = pdn_estatus.Trim();

                                    //cmd1.Connection = MyConnection;
                                    //cmd1.ExecuteNonQuery();
                                    //MyConnection.Close();
                                    //cmd1.Dispose();

                                    //    if (pdn_estatus.Trim() == "T")
                                    //    {
                                    //        if (Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) < pdn_num_unidades)
                                    //        {
                                    //            Fn_peso_unitario(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value));
                                    //            var_peso_unitario = var_dec_peso_total;
                                    //            Fn_invpt_salidas(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value), pdn_num_unidades, var_peso_unitario);                                                
                                    //        }
                                    //    }
                                    //}//if sino encontro en la tb_det_pedidos 
                                    reader2.Close();
                                }//for
                                //st.Stop();
                                //MessageBox.Show("Tiempo de proceso " + st.ElapsedMilliseconds);
                                //st.Reset();
                                #endregion

                                #region consulta la tb_det_pedidos para verificar que lo que hay en la tabla lo haya en el DGV2
                                /*cmnd2 = thisConnection.CreateCommand();
                                cmnd2.CommandText = "select lin_clave, prod_clave, pdn_tipo, pdn_num_unidades from tb_det_pedidos where pdn_folio = '" + txtfolio.Text + "'and pdn_tipo = 'EXP' order by nrenglon asc";
                                reader2 = cmnd2.ExecuteReader();
                                int j = 0;
                                while (reader2.Read())
                                {
                                    lin_clave = reader2.GetSqlString(0).ToString().Trim();
                                    prod_clave = reader2.GetSqlString(1).ToString().Trim();
                                    pdn_tipo = reader2.GetSqlString(2).ToString().Trim();
                                    pdn_num_unidades = reader2.GetSqlDecimal(3).Value;

                                    if (pdn_tipo.Trim() == "EXP")
                                    {
                                        string hay = "N";
                                        j = 0;
                                        while (j <= DGV2.Rows.Count - 1)
                                        {
                                            if (Convert.ToString(DGV2.Rows[j].Cells["linea"].Value).Trim() == lin_clave && Convert.ToString(DGV2.Rows[j].Cells["clave_producto"].Value).Trim() == prod_clave && Convert.ToDecimal(DGV2.Rows[j].Cells["cajas"].Value) > 0)
                                            {
                                                hay = "S";
                                                goto salir;
                                            }
                                            j++;
                                        }//while                                       
                                        if (hay == "N")
                                        {
                                            cmnd3 = thisConnection.CreateCommand();
                                            cmnd3.CommandText = "delete tb_det_pedidos where pdn_folio='" + txtfolio.Text + "' and prod_clave='" + prod_clave + "' and pdn_tipo = 'EXP'";
                                            reader3 = cmnd3.ExecuteReader();
                                            reader3.Close();
                                            goto salir;
                                        }                                            
                                    salir:
                                        if (pdn_estatus.Trim() == "T")
                                        {
                                            Fn_peso_unitario(lin_clave, prod_clave);
                                            var_peso_unitario = var_dec_peso_total;
                                            Fn_invpt_salidas(lin_clave, prod_clave, pdn_num_unidades, var_peso_unitario);
                                        }                                     
                                    }//si es nal 
                                }//while
                                reader2.Close();*/
                                #endregion

                                //cmnd2 = thisConnection.CreateCommand();
                                //cmnd2.CommandText = "select pdn_estatus from tb_mstr_pedidos_exp where pdn_folio ='" + txtfolio.Text + "'";
                                //reader2 = cmnd2.ExecuteReader();
                                //while (reader2.Read())
                                //{
                                //    pdn_estatus = reader2.GetSqlString(0).ToString();
                                //}
                                //reader2.Close();

                                if (pdn_estatus.Trim() == "")
                                {
                                    if (MessageBox.Show("¿Desea mandar a inventario la orden de venta?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                    {
                                        //MessageBox.Show("modifica inventario");
                                        //st.Start();
                                        cmnd2 = thisConnection.CreateCommand();
                                        cmnd2.CommandText = "update tb_mstr_pedidos_exp set pdn_estatus = 'T' where pdn_folio = '" + txtfolio.Text + "'";
                                        reader2 = cmnd2.ExecuteReader();
                                        reader2.Close();

                                        //MyConnection.Open();
                                        //dbCmdNull1 = MyConnection.CreateCommand();
                                        //dbCmdNull1.CommandText = "SET NULL OFF";
                                        //dbCmdNull1.ExecuteNonQuery();

                                        //cmd1.CommandText = "update tb_mstr_pedidos_nal set pdn_estatus = 'T' where pdn_folio = '" + txtfolio.Text + "' and pdn_tipo ='EXP'";                                        
                                        //cmd1.Connection = MyConnection;
                                        //cmd1.ExecuteNonQuery();
                                        //MyConnection.Close();
                                        //cmd1.Dispose();

                                        for (int i = 0; i < DGV2.Rows.Count; i++)
                                        {
                                            Fn_peso_unitario(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value));
                                            var_peso_unitario = var_dec_peso_total;

                                            Fn_invpt_salidas(Convert.ToString(DGV2.Rows[i].Cells["linea"].Value), Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value), Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value), var_peso_unitario);
                                        }//for                            
                                        //st.Stop();
                                        //MessageBox.Show("Tiempo en modificar inventario " + st.ElapsedMilliseconds);
                                        //st.Reset();
                                    }//enviar orden de venta a inventario
                                }//if estatus en blanco
                                thisConnection.Close();
                                //Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "M", "4.1.1", txtfolio.Text, "MODIFICACION ORDEN DE VENTA DE EXP CON EL FOLIO " + txtfolio.Text, "SIPGAB");
                                if (MessageBox.Show("¿Desea imprimir esta orden de venta?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                {
                                    imprimir(Utilerias.Class1.Nombre_impresora);
                                }
                                Utilerias.Class1.registrar_movimiento3(
                                    DateTime.Now,
                                    Environment.MachineName,
                                    Utilerias.Class1.Usu_login,
                                    "M",
                                    "4.1.1",
                                    txtfolio.Text,
                                    "MODIFICACION A LA O. VTA EXP F-" + txtfolio.Text + " Imp $" + txtusd.Text + " Moneda: " + CBmonfle.Text,
                                    "SIPGAB",
                                    txtfolio.Text,
                                    "Tot Gen $" + txtusd.Text + " VS Tot Det $ " + tot + " " + mProd);
                            }
                            catch (SqlException ex)
                            {
                                thisConnection.Close();
                                Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "sisgab", ex.ToString());
                                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                            catch (Exception ex1)
                            {
                                thisConnection.Close();
                                Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "sisgab", ex1.ToString());
                                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                        }
                        else
                        {
                            return;
                        }
                        if (Faltantes.Rows.Count > 0)
                        {
                            Frm_Faltante ProdFal = new Frm_Faltante();
                            ProdFal.datos = Faltantes;
                            Program.MyGlobal.PubPed = txtfolio.Text;
                            Program.MyGlobal.PubCnte = txtcliente.Text;
                            Program.MyGlobal.PubNomCnte = CBCliente.Text.Trim();
                            Program.MyGlobal.PubFec = txtfecha.Text.ToString();
                            ProdFal.muser = Utilerias.Class1.Usu_login;
                            ProdFal.ShowDialog(this);
                        }
                        //}
                    }
                    break;
                #endregion

                #region volver a imprimir orden de venta
                case 4:
                    if (MessageBox.Show("¿Desea Imprimir?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        imprimir(Utilerias.Class1.Nombre_impresora);
                        //imprimir();
                    }
                    break;
                    #endregion

            }//switch
            btnCancel_Click(sender, e);
        }

        //se determina el peso real del producto
        private void Fn_peso_unitario(string prm_lin_clave, string prm_prod_clave)
        {
            string um_clave;
            decimal compp_peso, um_equivalencia = 0;
            cmnd3 = thisConnection.CreateCommand();
            cmnd3.CommandText = "select um_clave, compp_peso from tb_mstr_comp_prod where lin_clave= '" + prm_lin_clave + "'and prod_clave = '" + prm_prod_clave + "'";
            reader3 = cmnd3.ExecuteReader();
            while (reader3.Read())
            {
                um_clave = reader3.GetSqlString(0).ToString();
                compp_peso = reader3.GetSqlDecimal(1).Value;

                cmnd4 = thisConnection.CreateCommand();
                cmnd4.CommandText = "select um_equivalencia from tb_cat_unidad where um_clave = '" + um_clave + "'";
                reader4 = cmnd4.ExecuteReader();
                while (reader4.Read())
                {
                    um_equivalencia = reader4.GetSqlDecimal(0).Value;
                }
                if (reader4.HasRows == false)
                {
                    um_equivalencia = 1;
                }
                reader4.Close();

                var_dec_peso_total = var_dec_peso_total + (compp_peso * um_equivalencia);
            }
            reader3.Close();
        }

        //funcion que resta las unidades en el inventario de pt
        private void Fn_invpt_salidas(string prm_lin_clave, string prm_prod_clave, decimal prm_num_unidades, decimal prm_peso_unitario)
        {
            //MessageBox.Show("Inicio conteo");
            Stopwatch stopw = new Stopwatch();
            stopw.Start();
            //for (int rows = 0; rows < DGV2.Rows.Count; rows++)
            //{

            // INVENTARIO PT CUANDO NUNCA SE HA DADO DE ALTA EL PRODUCTO
            string varfechafox = DateTime.Now.Date.ToString("MM/dd/yyyy");
            string varfecha = DateTime.Now.Date.ToString("dd/MM/yyyy");
            SqlCommand cmndInvPT1 = thisConnection.CreateCommand();
            SqlDataReader readerInvPT1;
            cmndInvPT1.CommandText = "SELECT * FROM  tb_mstr_inventario_pt WHERE (prod_clave = '" + prm_prod_clave + "')";
            string encontrado = "N";
            readerInvPT1 = cmndInvPT1.ExecuteReader();
            while (readerInvPT1.Read())
            {
                encontrado = "S";
                break;
            }
            cmndInvPT1.Dispose();
            readerInvPT1.Dispose();
            if (encontrado.Equals("N")) // INVENTARIO PT SI NO SE ENCONTRO EL PRODUCTO SE DA DE ALTA NORMALMENTE
            {
                string insertarSql3 = "INSERT INTO tb_mstr_inventario_pt(invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, " +
                                            "invpt_entradas_un, invpt_entradas_kg, invpt_salidas_un, invpt_salidas_kg) " +
                              "VALUES " +
                              "(@aux_invpt_fecha,@aux_lin_clave, @aux_prod_clave, @aux_invpt_inicial_un, @aux_invpt_inicial_kg, " +
                              "@aux_invpt_entradas_un, @aux_invpt_entradas_kg, @aux_invpt_salidas_un, @aux_invpt_salidas_kg)";
                SqlCommand insertarCmd3 = new SqlCommand(insertarSql3, thisConnection);
                insertarCmd3.Parameters.AddWithValue("@aux_invpt_fecha", DateTime.Now.ToString("s"));
                insertarCmd3.Parameters.AddWithValue("@aux_lin_clave", prm_lin_clave);
                insertarCmd3.Parameters.AddWithValue("@aux_prod_clave", prm_prod_clave);
                insertarCmd3.Parameters.AddWithValue("@aux_invpt_inicial_un", 0);
                insertarCmd3.Parameters.AddWithValue("@aux_invpt_inicial_kg", 0.0);
                insertarCmd3.Parameters.AddWithValue("@aux_invpt_entradas_un", prm_num_unidades);
                insertarCmd3.Parameters.AddWithValue("@aux_invpt_entradas_kg", var_peso_unitario);
                insertarCmd3.Parameters.AddWithValue("@aux_invpt_salidas_un", 0);
                insertarCmd3.Parameters.AddWithValue("@aux_invpt_salidas_kg", 0.0);
                insertarCmd3.ExecuteNonQuery();

                //MessageBox.Show("Graba cuando no lo encuentra");
                //MyConnection.Open();
                //OleDbCommand dbCmdNull6 = MyConnection.CreateCommand();
                //dbCmdNull6.CommandText = "SET NULL OFF";
                //dbCmdNull6.ExecuteNonQuery();
                //String aux_invpt_fecha = DateTime.Now.ToString().Substring(0, 10);
                //OleDbCommand cmd6 = new OleDbCommand();
                //cmd6.CommandType = CommandType.Text;
                //cmd6.CommandText = "INSERT INTO tb_mstr_inventario_pt ([INVPT_FECHA],[LIN_CLAVE],[PROD_CLAVE],[INVPT_INICIAL_UN],[INVPT_INICIAL_KG],[INVPT_ENTRADAS_UN],[INVPT_ENTRADAS_KG],[INVPT_SALIDAS_UN],[INVPT_SALIDAS_KG]) VALUES (?,?,?,?,?,?,?,?,?)";
                //cmd6.Parameters.Add("@Nombre0", OleDbType.DBDate).Value = aux_invpt_fecha;
                //cmd6.Parameters.Add("@Nombre1", OleDbType.Char).Value = prm_lin_clave;
                //cmd6.Parameters.Add("@Nombre2", OleDbType.Char).Value = prm_prod_clave;
                //cmd6.Parameters.Add("@Nombre3", OleDbType.Numeric).Value = 0;
                //cmd6.Parameters.Add("@Nombre4", OleDbType.Numeric).Value = 0.0;
                //cmd6.Parameters.Add("@Nombre5", OleDbType.Numeric).Value = prm_num_unidades;
                //cmd6.Parameters.Add("@Nombre6", OleDbType.Numeric).Value = var_peso_unitario;
                //cmd6.Parameters.Add("@Nombre7", OleDbType.Numeric).Value = 0;
                //cmd6.Parameters.Add("@Nombre8", OleDbType.Numeric).Value = 0.0;
                //cmd6.Connection = MyConnection;
                //cmd6.ExecuteNonQuery();
                //MyConnection.Close();
                ////MessageBox.Show("ya grabo cuando no lo encontro");
            }
            else
            {
                // INVENTARIO PT SI NO SE ENCONTRO EL PRODUCTO BUSCAMOS EN EL DIA ACTUAL JUNTO CON LA CLAVE DEL PRODUCTO
                SqlCommand cmndInvPT2 = thisConnection.CreateCommand();
                SqlDataReader readerInvPT2;
                cmndInvPT2.CommandText = "SELECT * FROM  tb_mstr_inventario_pt WHERE (invpt_fecha = '" + varfecha + "') AND (prod_clave = '" + prm_prod_clave + "')";
                encontrado = "N";
                readerInvPT2 = cmndInvPT2.ExecuteReader();
                while (readerInvPT2.Read())
                {
                    encontrado = "S";
                    break;
                }
                cmndInvPT2.Dispose();
                readerInvPT2.Dispose();
                if (encontrado.Equals("S"))
                {
                    // INVENTARIO PT SI SE ENCONTRO EL PRODUCTO EN LA FECHA ACTUAL DE LA CAPTURA DE PT ENTONCES SOLO ACTUALIZAMOS DATOS
                    string SQLInvPT = "UPDATE tb_mstr_inventario_pt SET invpt_entradas_un = (invpt_entradas_un + " + prm_num_unidades + "), invpt_entradas_kg = (invpt_entradas_kg + " + var_peso_unitario + ") " +
                                        "WHERE  (invpt_fecha = '" + varfecha + "') AND (prod_clave = '" + prm_prod_clave + "')";
                    SqlCommand cmdInvPT = new SqlCommand(SQLInvPT, thisConnection);
                    cmdInvPT.ExecuteNonQuery();
                    cmdInvPT.Dispose();

                    //MessageBox.Show("actualiza cuando lo encuentra con la fecha exacta");
                    ////MyConnection.Open();   // CONEXION A VISUAL FOX
                    ////OleDbCommand dbCmdNull6 = MyConnection.CreateCommand();
                    ////dbCmdNull6.CommandText = "SET NULL OFF";
                    ////dbCmdNull6.ExecuteNonQuery();
                    ////OleDbCommand cmd6 = new OleDbCommand();
                    ////cmd6.CommandType = CommandType.Text;

                    ////cmd6.CommandText = "UPDATE tb_mstr_inventario_pt SET INVPT_INICIAL_UN = (INVPT_INICIAL_UN + " + prm_num_unidades + "), INVPT_INICIAL_KG = (INVPT_INICIAL_KG + " + var_peso_unitario + ") " +
                    ////                   "WHERE  (INVPT_FECHA  BETWEEN {" + varfechafox + "} AND {" + varfechafox + "}) AND (PROD_CLAVE = '" + prm_prod_clave + "')";


                    ////cmd6.Connection = MyConnection;
                    ////cmd6.ExecuteNonQuery();
                    ////MyConnection.Close();
                    //////MessageBox.Show("ya actualizo cuando encontro la fecha exacta");
                }
                else
                {
                    // INVENTARIO PT EN CASO QUE NO SE ENCONTRO ESE PRODUCTO EN EL DIA ACTUAL BUSCAMOS LA ULTIMA CAPTURA DE ESE PRODUCTO PARA TRAERLO COMO INVENTARIO ACTUAL Y ACTUALIZAMOS TODOS LOS DATOS DE UNIDADES Y KGS
                    Double valor_un_invanterior = 0; Double valor_un_entanterior = 0; Double valor_un_salanterior = 0;
                    Double valor_kg_invanterior = 0; Double valor_kg_entanterior = 0; Double valor_kg_salanterior = 0;
                    Double total_un_invanterior = 0;
                    Double total_kg_invanterior = 0;
                    SqlCommand cmndInvPT3 = thisConnection.CreateCommand();
                    SqlDataReader readerInvPT3;
                    cmndInvPT3.CommandText = "SELECT TOP (1) isnull(invpt_inicial_un, 0), isnull(invpt_entradas_un, 0), isnull(invpt_salidas_un, 0), isnull(invpt_inicial_kg, 0), " +
                                             "isnull(invpt_entradas_kg, 0), isnull(invpt_salidas_kg, 0) FROM  tb_mstr_inventario_pt WHERE (prod_clave = '" + prm_prod_clave + "') ORDER BY invpt_fecha DESC";
                    readerInvPT3 = cmndInvPT3.ExecuteReader();
                    while (readerInvPT3.Read())
                    {
                        valor_un_invanterior = Convert.ToDouble(readerInvPT3.GetSqlDecimal(0).Value);
                        valor_un_entanterior = Convert.ToDouble(readerInvPT3.GetSqlDecimal(1).Value);
                        valor_un_salanterior = Convert.ToDouble(readerInvPT3.GetSqlDecimal(2).Value);

                        valor_kg_invanterior = Convert.ToDouble(readerInvPT3.GetSqlDecimal(3).Value);
                        valor_kg_entanterior = Convert.ToDouble(readerInvPT3.GetSqlDecimal(4).Value);
                        valor_kg_salanterior = Convert.ToDouble(readerInvPT3.GetSqlDecimal(5).Value);

                        encontrado = "S";
                    }
                    cmndInvPT1.Dispose();
                    readerInvPT1.Dispose();
                    total_un_invanterior = (valor_un_invanterior + valor_un_entanterior) - valor_un_salanterior;
                    total_kg_invanterior = (valor_kg_invanterior + valor_kg_entanterior) - valor_kg_salanterior;

                    string insertarSql3 = "INSERT INTO tb_mstr_inventario_pt(invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, " +
                                                "invpt_entradas_un, invpt_entradas_kg, invpt_salidas_un, invpt_salidas_kg) " +
                                  "VALUES " +
                                  "(@aux_invpt_fecha,@aux_lin_clave, @aux_prod_clave, @aux_invpt_inicial_un, @aux_invpt_inicial_kg, " +
                                  "@aux_invpt_entradas_un, @aux_invpt_entradas_kg, @aux_invpt_salidas_un, @aux_invpt_salidas_kg)";
                    SqlCommand insertarCmd3 = new SqlCommand(insertarSql3, thisConnection);
                    insertarCmd3.Parameters.AddWithValue("@aux_invpt_fecha", DateTime.Now.ToString("s"));
                    insertarCmd3.Parameters.AddWithValue("@aux_lin_clave", prm_lin_clave);
                    insertarCmd3.Parameters.AddWithValue("@aux_prod_clave", prm_prod_clave);
                    insertarCmd3.Parameters.AddWithValue("@aux_invpt_inicial_un", total_un_invanterior);
                    insertarCmd3.Parameters.AddWithValue("@aux_invpt_inicial_kg", total_kg_invanterior);
                    insertarCmd3.Parameters.AddWithValue("@aux_invpt_entradas_un", prm_num_unidades);
                    insertarCmd3.Parameters.AddWithValue("@aux_invpt_entradas_kg", var_peso_unitario);
                    insertarCmd3.Parameters.AddWithValue("@aux_invpt_salidas_un", 0);
                    insertarCmd3.Parameters.AddWithValue("@aux_invpt_salidas_kg", 0.0);
                    insertarCmd3.ExecuteNonQuery();

                    //MessageBox.Show("inserta cuando saca el ultimo registro");
                    //MyConnection.Open();   // CONEXION A VISUAL FOX
                    //OleDbCommand dbCmdNull6 = MyConnection.CreateCommand();
                    //dbCmdNull6.CommandText = "SET NULL OFF";
                    //dbCmdNull6.ExecuteNonQuery();
                    //String aux_invpt_fecha = DateTime.Now.ToString().Substring(0, 10);
                    //OleDbCommand cmd6 = new OleDbCommand();
                    //cmd6.CommandType = CommandType.Text;
                    //cmd6.CommandText = "INSERT INTO tb_mstr_inventario_pt ([INVPT_FECHA],[LIN_CLAVE],[PROD_CLAVE],[INVPT_INICIAL_UN],[INVPT_INICIAL_KG],[INVPT_ENTRADAS_UN],[INVPT_ENTRADAS_KG],[INVPT_SALIDAS_UN],[INVPT_SALIDAS_KG]) VALUES (?,?,?,?,?,?,?,?,?)";
                    //cmd6.Parameters.Add("@Nombre0", OleDbType.DBDate).Value = aux_invpt_fecha;
                    //cmd6.Parameters.Add("@Nombre1", OleDbType.Char).Value = prm_lin_clave;
                    //cmd6.Parameters.Add("@Nombre2", OleDbType.Char).Value = prm_prod_clave;
                    //cmd6.Parameters.Add("@Nombre3", OleDbType.Numeric).Value = total_un_invanterior;
                    //cmd6.Parameters.Add("@Nombre4", OleDbType.Numeric).Value = total_kg_invanterior;
                    //cmd6.Parameters.Add("@Nombre5", OleDbType.Numeric).Value = prm_num_unidades;
                    //cmd6.Parameters.Add("@Nombre6", OleDbType.Numeric).Value = var_peso_unitario;
                    //cmd6.Parameters.Add("@Nombre7", OleDbType.Numeric).Value = 0;
                    //cmd6.Parameters.Add("@Nombre8", OleDbType.Numeric).Value = 0.0;
                    //cmd6.Connection = MyConnection;
                    //cmd6.ExecuteNonQuery();
                    //MyConnection.Close();
                    ////MessageBox.Show("ya grabo cuando saco el ultimo registro");
                }
            }
            //}//for
            stopw.Stop();
            //MessageBox.Show("tiempo trancurrido en modificar el inventario de pt: " + stopw.ElapsedMilliseconds);       
        }

        //funcio que sum las unidades al inventario de pt
        private void Fn_invpt_entradas(string prm_lin_clave, string prm_prod_clave, decimal prm_num_unidades, decimal prm_peso_unitario)
        {
            Stopwatch sw = new Stopwatch();
            //MessageBox.Show("Inicio conteo entradas");
            sw.Start();
            decimal invpt_inicial_kg = 0, invpt_entradas_kg = 0, invpt_salidas_kg = 0, invpt_inicial_un = 0, invpt_entradas_un = 0, invpt_salidas_un = 0, var_dec_inv_inicial_kg = 0, var_dec_inv_inicial_un = 0;
            cmnd3 = thisConnection.CreateCommand();
            cmnd3.CommandText = "select isnull(invpt_entradas_kg,0) as invpt_entradas_kg, isnull(invpt_entradas_un,0) as invpt_entradas_un from tb_mstr_inventario_pt " +
                                "where invpt_fecha='" + DateTime.Now.ToShortDateString() + "' and lin_clave = '" + prm_lin_clave + "' and prod_clave ='" + prm_prod_clave + "'";
            reader3 = cmnd3.ExecuteReader();
            while (reader3.Read())
            {
                invpt_entradas_kg = invpt_entradas_kg + (prm_peso_unitario * prm_num_unidades);
                invpt_entradas_un = invpt_entradas_un + prm_num_unidades;

                cmnd4 = thisConnection.CreateCommand();
                cmnd4.CommandText = "update tb_mstr_inventario_pt set invpt_entradas_kg = '" + invpt_entradas_kg + "', invpt_entradas_un = '" + invpt_entradas_un + "' where lin_clave='" + prm_lin_clave + "' and prod_clave = '" + prm_prod_clave + "' and invpt_fecha='" + DateTime.Now.ToString("s") + "'";
                reader4 = cmnd4.ExecuteReader();
                reader4.Close();

                //string fec = DateTime.Now.ToString("s");
                //cmd1.Dispose();
                //MyConnection.Open();
                //cmd1.CommandType = CommandType.Text;
                //cmd1.CommandText = "update tb_mstr_inventario_pt set invpt_entradas_un = ? where lin_clave='" + prm_lin_clave + "' and prod_clave = '" + prm_prod_clave + "' and (invpt_fecha between {" + fec + "} and {" + fec + "})";
                //cmd1.Parameters.Add("@Nombre100", OleDbType.Numeric);
                //cmd1.Parameters["@Nombre100"].Value = invpt_entradas_kg;
                //cmd1.Parameters.Add("@Nombre101", OleDbType.Numeric, 6, "invpt_entradas_un");
                //cmd1.Parameters["@Nombre101"].Value = invpt_entradas_un;
                //cmd1.Connection = MyConnection;
                //cmd1.ExecuteNonQuery();
                //MyConnection.Close();
                //cmd1.Dispose();
            }
            if (reader3.HasRows == false)
            {
                cmnd4 = thisConnection.CreateCommand();
                cmnd4.CommandText = "SELECT TOP 1 isnull(invpt_inicial_kg,0) as invpt_inicial_kg, isnull(invpt_entradas_kg,0) as invpt_entradas_kg, " +
                                    "isnull(invpt_salidas_kg,0) as invpt_salidas_kg, isnull(invpt_inicial_un,0) as invpt_inicial_un, isnull(invpt_entradas_un,0) as invpt_entradas_un, " +
                                    "isnull(invpt_salidas_un,0) as invpt_salidas_un FROM tb_mstr_inventario_pt WHERE lin_clave='" + prm_lin_clave.Trim() + "' AND prod_clave= '" + prm_prod_clave + "' ORDER BY invpt_fecha DESC";
                reader4 = cmnd4.ExecuteReader();
                while (reader4.Read())
                {
                    invpt_inicial_kg = reader4.GetSqlDecimal(0).Value;
                    invpt_entradas_kg = reader4.GetSqlDecimal(1).Value;
                    invpt_salidas_kg = reader4.GetSqlDecimal(2).Value;
                    invpt_inicial_un = reader4.GetSqlDecimal(3).Value;
                    invpt_entradas_un = reader4.GetSqlDecimal(4).Value;
                    invpt_salidas_un = reader4.GetSqlDecimal(5).Value;

                    var_dec_inv_inicial_kg = invpt_inicial_kg + invpt_entradas_kg - invpt_salidas_kg;
                    var_dec_inv_inicial_un = invpt_inicial_un + invpt_entradas_un - invpt_salidas_un;
                    invpt_salidas_kg = prm_peso_unitario * prm_num_unidades;

                    cmnd6 = thisConnection.CreateCommand();
                    cmnd6.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_salidas_kg, invpt_salidas_un, invpt_entradas_kg, invpt_entradas_un) values ('" + DateTime.Now.ToString("s") + "','" + prm_lin_clave + "','" + prm_prod_clave + "','" + var_dec_inv_inicial_un + "','" + var_dec_inv_inicial_kg + "','" + invpt_salidas_kg + "','" + prm_num_unidades + "', '0', '0')";
                    reader6 = cmnd6.ExecuteReader();
                    reader6.Close();

                    //cmd1.Dispose();
                    //MyConnection.Open();
                    //cmd1.CommandType = CommandType.Text;
                    //OleDbCommand dbCmdNull1 = MyConnection.CreateCommand();
                    //dbCmdNull1.CommandText = "SET NULL OFF";
                    //dbCmdNull1.ExecuteNonQuery();

                    //string fec = DateTime.Now.ToString("s");
                    //cmd1.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_salidas_kg, invpt_salidas_un, invpt_entradas_kg, invpt_entradas_un) values (?,?,?,?,?,?,?,?,?)";
                    //cmd1.Parameters.Add("@Nombre0", OleDbType.Date);
                    //cmd1.Parameters["@Nombre0"].Value = fec;
                    //cmd1.Parameters.Add("@Nombre1", OleDbType.Char);
                    //cmd1.Parameters["@Nombre1"].Value = prm_lin_clave;
                    //cmd1.Parameters.Add("@Nombre2", OleDbType.Char);
                    //cmd1.Parameters["@Nombre2"].Value = prm_prod_clave;
                    //cmd1.Parameters.Add("@Nombre3", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre3"].Value = var_dec_inv_inicial_un;
                    //cmd1.Parameters.Add("@Nombre4", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre4"].Value = var_dec_inv_inicial_kg;
                    //cmd1.Parameters.Add("@Nombre5", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre5"].Value = invpt_salidas_kg;
                    //cmd1.Parameters.Add("@Nombre6", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre6"].Value = prm_num_unidades;
                    //cmd1.Parameters.Add("@Nombre7", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre7"].Value = 0;
                    //cmd1.Parameters.Add("@Nombre8", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre8"].Value = 0;

                    //cmd1.Connection = MyConnection;
                    //cmd1.ExecuteNonQuery();
                    //MyConnection.Close();
                    //cmd1.Dispose();
                }
                if (reader4.HasRows == false)
                {
                    invpt_salidas_kg = (prm_peso_unitario * prm_num_unidades);
                    cmnd6 = thisConnection.CreateCommand();
                    cmnd6.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_salidas_un, invpt_salidas_kg, invpt_entradas_kg, invpt_entradas_un) values ('" + DateTime.Now.ToString("s") + "','" + prm_lin_clave + "','" + prm_prod_clave + "', 0, 0, '" + invpt_salidas_kg + "','" + prm_num_unidades + "', '0', '0')";
                    reader6 = cmnd6.ExecuteReader();
                    reader6.Close();

                    //MyConnection.Open();
                    //OleDbCommand cmd1 = new OleDbCommand();
                    //cmd1.CommandType = CommandType.Text;
                    //OleDbCommand dbCmdNull1 = MyConnection.CreateCommand();
                    //dbCmdNull1.CommandText = "SET NULL OFF";
                    //dbCmdNull1.ExecuteNonQuery();

                    //string fec = DateTime.Now.ToString("s");
                    //cmd1.CommandText = "insert into tb_mstr_inventario_pt (invpt_fecha, lin_clave, prod_clave, invpt_inicial_un, invpt_inicial_kg, invpt_salidas_un, invpt_salidas_kg, invpt_entradas_kg, invpt_entradas_un) values (?,?,?,?,?,?,?,?,?)";
                    //cmd1.Parameters.Add("@Nombre0", OleDbType.Date);
                    //cmd1.Parameters["@Nombre0"].Value = fec;
                    //cmd1.Parameters.Add("@Nombre1", OleDbType.Char);
                    //cmd1.Parameters["@Nombre1"].Value = prm_lin_clave;
                    //cmd1.Parameters.Add("@Nombre2", OleDbType.Char);
                    //cmd1.Parameters["@Nombre2"].Value = prm_prod_clave;
                    //cmd1.Parameters.Add("@Nombre3", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre3"].Value = 0;
                    //cmd1.Parameters.Add("@Nombre4", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre4"].Value = 0;
                    //cmd1.Parameters.Add("@Nombre5", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre5"].Value = invpt_salidas_kg;
                    //cmd1.Parameters.Add("@Nombre6", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre6"].Value = prm_num_unidades;
                    //cmd1.Parameters.Add("@Nombre7", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre7"].Value = 0;
                    //cmd1.Parameters.Add("@Nombre8", OleDbType.Numeric);
                    //cmd1.Parameters["@Nombre8"].Value = 0;

                    //cmd1.Connection = MyConnection;
                    //cmd1.ExecuteNonQuery();
                    //MyConnection.Close();
                    //cmd1.Dispose();
                }
                reader4.Close();
            }
            reader3.Close();
            sw.Stop();
            //MessageBox.Show("Fin del proceso de entradas, tiempo transcurrido: " + sw.ElapsedMilliseconds);
        }

        //método para imprimir
        private void imprimir(string impresora)
        {
            //printDialog1.Document = new PrintDocument();
            //PrintPreviewDialog VistaPrevia = new PrintPreviewDialog();
            //printDialog1.Document.PrintPage += printDocument1_PrintPage;
            //VistaPrevia.Document = printDialog1.Document;
            //VistaPrevia.ShowDialog();

            PrintDocument pd = new PrintDocument();

            pd.PrintPage += new PrintPageEventHandler(this.printDocument1_PrintPage);
            //if (Utilerias.Class1.Nombre_impresora == null)
            //{
            //Utilerias.Class1.Nombre_impresora = "Canon iR2220/iR3320 PCL6";
            //    //Utilerias.Nombre_impresora = "Foxit Reader PDF Printer";
            //}
            //else
            //{
            // Original   
            Boolean Existe = false;
            foreach (String strPrinter in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {

                //if (strPrinter.Contains("Canon"))
                //{
                //    pd.PrinterSettings.PrinterName = strPrinter;
                //    Existe = true;
                //    break;
                //}
            }
            if (!Existe)
                pd.PrinterSettings.PrinterName = Utilerias.Class1.Nombre_impresora;
            //pd.PrinterSettings.PrinterName = pd.PrinterSettings.d;

            // para pruebas
            //pd.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
            //
            //}
            pd.Print();
        }

        //valida txtbox
        private void validar()
        {
            if (txtobs.Text == string.Empty)
                txtobs.Text = " ";

            if (txtpedsigma.Text == string.Empty)
                txtpedsigma.Text = " ";

            if (txttalonemb.Text == string.Empty)
                txttalonemb.Text = " ";

            if (txtimporte.Text == string.Empty)
                txtimporte.Text = "0";

            if (txtcajatrailer.Text == string.Empty)
                txtcajatrailer.Text = " ";

            if (txtsello.Text == string.Empty)
                txtsello.Text = " ";

            if (txtimporte.Text == string.Empty)
                txtimporte.Text = "0";
        }


        //imprime la orden de venta
        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            sumataraprox = sumacomp = sumasplit = decimales = enteros = 0;
            sumapedido = 0;
            String drawString1 = " ", drawString2 = " ", drawString3 = " ", drawString4 = " ", drawobs = " ", drawryan = " ";
            String drawLinea = " ", drawProd = " ", drawPedi = " ", sumatar = " ", taraprox = " ", sumcomp = " ", sumdeci = " ", drawcajas = " ", drawtaraprox = " ", drawobsprod = " ";
            String drawstring8 = " ", drawstring9 = " ", drawstring10 = " ", drawstring11 = " ", drawstring12 = " ";
            decimal Totalpeso = 0;
            DateTime dt = DateTime.Now;
            //int cont = 148, cont1 = 165;
            int cont = 148, cont1 = 160, cont2 = 160;
            cont_prod = 1;
            CreaQr(txtfolio.Text.Trim());
            //thisConnection.Open();
            String drawString = "";
            if (ultimo_renglon == 30)
            {
                drawString = "\t\t\t\t\tCONTINUACION";
                //drawString2 = "_________________________________________________________________________________________________________________________________________";
                drawString3 = "";
                drawString4 = "";
                drawstring9 = "";
                drawstring10 = "";
                drawstring11 = "";
            }
            else
            {
                drawString = "\t\t\t\t\tComercializadora GAB, S.A. de C.V.\r\n " +
                                "\t\t\t\tORDEN DE VENTA EXPORTACION/REPORTE DE EMBARQUES\r\n\r\n";
                //drawString1 = "*" + txtfolio.Text + "*\r\n";
                drawString1 = "*" + txtfolio.Text + "*\r\n";
                drawString3 = "\tCLIENTE: " + txtcliente.Text.Trim() + " " + CBCliente.SelectedItem.ToString().Trim() + "";
                drawString4 = "\t\tFOLIO: " + txtfolio.Text + "\t FECHA: " + txtfecha.Text + "\r\n";
                drawstring9 = "HORA DE EMISION: " + dt.ToString("T").Trim() + "\r\n";
                drawstring10 = "\tTRANSPORTISTA: " + CBTransportista.SelectedItem.ToString().Trim() + "\t CARTA PORTE: " + txttalonemb.Text.Trim() + "";
                drawstring11 = "IRAPUATO GTO.\r\n\r\n";
            }
            drawString2 = "_________________________________________________________________________________________________________________________________________";
            //drawString1 = "*" + txtfolio.Text + "*\r\n";


            drawstring12 = "\tDETALLE DE LA ORDEN DE VENTA\r\n" +
                                "Linea     Producto  \t\t\t\t\t\t\t\t        |  Pedido  |  Tarima    \t\t\t|  Tar. Aprox.  | Kg Caja | Peso Tot \r\n\r\n";

            // Create font and brush. 
            Font drawFont = new Font("Arial", 8);//encabezado            
            Font drawFont1 = new Font("PF Barcode 39", 25);//código de barras
            Font drawFont3 = new Font("Arial", 5);//tarima
            SolidBrush drawBrush = new SolidBrush(Color.Black);

            // Create point for upper-left corner of drawing. 
            PointF drawPoint = new PointF(10.0F, 10.0F);//encabezado
            PointF drawPoint1 = new PointF(620.0F, 20.0F);//codigo de barras            
            PointF drawPoint2 = new PointF(00.0f, 42.0f);//linea1
            PointF drawPoint3 = new PointF(0.0f, 55.0f);//datos1
            PointF drawpoint4 = new PointF(450.0F, 55.0F);
            PointF drawpoint5 = new PointF(0.0F, 68.0F);
            PointF drawpoint6 = new PointF(540.0F, 68.0F);
            PointF drawpoint7 = new PointF(0.0F, 82.0F);
            PointF drawpoint8 = new PointF(540.0F, 82.0F);
            PointF drawpoint9 = new PointF(0.0F, 105.0F);
            PointF drawPoint4 = new PointF(00.0f, 90.0f);//linea2
            PointF drawPoint5 = new PointF(00.0f, 130.0f);//linea3

            if (File.Exists(@"C:\reportes\qrcode.png"))
            {
                Image newImage = Image.FromFile(@"C:\reportes\qrcode.png"); //codigo para mandar llamar la imagen
                e.Graphics.DrawImage(newImage, 10, 10);
            }
            e.Graphics.DrawString(drawString1, drawFont1, drawBrush, drawPoint1);//código de barras            
            e.Graphics.DrawString(drawString, drawFont, drawBrush, drawPoint);//encabezado            
            e.Graphics.DrawString(drawString2, drawFont, drawBrush, drawPoint2);//linea1
            e.Graphics.DrawString(drawString3, drawFont, drawBrush, drawPoint3);//datos1
            e.Graphics.DrawString(drawString4, drawFont, drawBrush, drawpoint4);//
            e.Graphics.DrawString(drawstring8, drawFont, drawBrush, drawpoint5);//
            e.Graphics.DrawString(drawstring9, drawFont, drawBrush, drawpoint6);//
            e.Graphics.DrawString(drawstring10, drawFont, drawBrush, drawpoint7);//
            e.Graphics.DrawString(drawstring11, drawFont, drawBrush, drawpoint8);//
            e.Graphics.DrawString(drawstring12, drawFont, drawBrush, drawpoint9);//
            e.Graphics.DrawString(drawString2, drawFont, drawBrush, drawPoint4);//linea2
            e.Graphics.DrawString(drawString2, drawFont, drawBrush, drawPoint5);//linea3                       

            #region
            thisConnection.Open();
            //for (int i = 0; i < DGV2.Rows.Count; i++)
            for (int i = ultimo_renglon; i < DGV2.Rows.Count; i++)
            {
                if (cont_prod < 29)
                {
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select prod_num_tarimas from tb_cat_producto where prod_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "'";

                    reader2 = cmnd2.ExecuteReader();
                    while (reader2.Read())
                    {
                        prod_num_tarimas = reader2.GetSqlDecimal(0).Value;
                    }
                    reader2.Close();
                    if (prod_num_tarimas == 0)
                    {
                        resulta = 0;
                    }
                    else
                    {
                        resulta = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) / prod_num_tarimas;
                    }


                    drawLinea = string.Format("{0}", Convert.ToString(DGV2.Rows[i].Cells["linea"].Value));
                    if (Convert.ToString(DGV2.Rows[i].Cells["linea"].Value).Trim() == "08")
                    {
                        Font drawFont2 = new Font("Arial", 11);//encabezado
                        String modi1 = "PONER TARIMAS INTERMEDIAS DE SOPORTE Y AJUSTAR LAS 5 GATAS";
                        PointF drawmodi1 = new PointF(45.0f, 970.0f);//modificación 16/06/2014
                        e.Graphics.DrawString(modi1, drawFont2, drawBrush, drawmodi1);
                    }

                    if (cbdestino.Items.Count > 0)
                    {
                        if (Convert.ToString(cbdestino.SelectedItem).Trim() == "MCALLEN, TX.")
                        {
                            if ((Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("APIO")) || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("LECHUGA"))
                                || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("KALE")) || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("SWISS CHARD"))
                                || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("RED SWISS CHARD")) || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("CILANTRO"))
                                || (Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value).Contains("RED KALE")))
                            {
                                leyenda1 = "ASEGURARSE QUE EN ESTE EMBARQUE HAYA UNA DISTANCIA DE 18 PULGADAS ENTRE LA PARTE ALTA DEL PALLET Y EL TECHO PARA LOS QUE CRUZAN POR " + cbdestino.SelectedItem.ToString().Trim();
                                Font drawFont2 = new Font("Arial", 9);//encabezado                        
                                PointF drawmodi1 = new PointF(45.0f, 970.0f);//modificación 16/06/2014
                                e.Graphics.DrawString(leyenda1.Substring(0, 93), drawFont2, drawBrush, drawmodi1);

                                PointF drawmodi11 = new PointF(45.0f, 983.0f);//modificación 16/06/2014
                                e.Graphics.DrawString(leyenda1.Substring(93), drawFont2, drawBrush, drawmodi11);
                            }
                        }
                    }
                    drawProd = string.Format("{0}\t{1}", Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value), Convert.ToString(DGV2.Rows[i].Cells["nombre"].Value));
                    drawPedi = string.Format("|               |                     \t\t      |  \t\t|  \t\t\t\r\n");
                    drawcajas = (Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value)).ToString();
                    drawtaraprox = resulta.ToString("0.00");
                    if (Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value).ToString().Trim().Length > 30)
                    {
                        String cadena = Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value).Substring(0, 30);
                        PointF drawPointobs = new PointF(620.0F, cont);
                        String cadena2 = Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value).Substring(30);
                        PointF drawPointobs2 = new PointF(620.0F, cont + 13);
                        e.Graphics.DrawString(cadena, drawFont, drawBrush, drawPointobs);
                        e.Graphics.DrawString(cadena2, drawFont, drawBrush, drawPointobs2);
                    }
                    else
                    {
                        drawobsprod = Convert.ToString(DGV2.Rows[i].Cells["observaciones"].Value);
                        PointF drawPointobs = new PointF(620.0F, cont);
                        e.Graphics.DrawString(drawobsprod, drawFont, drawBrush, drawPointobs);
                    }

                    String tarima = string.Format("{0,15}", Convert.ToString(DGV2.Rows[i].Cells["nombre_tarima"].Value));
                    String subcli = string.Format("{0,10}", Convert.ToString(DGV2.Rows[i].Cells["subcnte_clave"].Value));
                    //String aduan = "Aduana: " + Convert.ToString(DGV2.Rows[i].Cells["aduana"].Value);

                    PointF drawPointlin = new PointF(0.0F, cont);
                    PointF drawPointProd = new PointF(30.0f, cont);
                    PointF drawPointPedi = new PointF(470.0F, cont);
                    PointF drawPointcajas = new PointF(490.0F, cont);
                    PointF drawtar = new PointF(523.0F, cont);
                    PointF drawPointtar = new PointF(680.0F, cont);
                    PointF drawsubcli = new PointF(523.0F, cont + 13);
                    PointF drawPoinPes = new PointF(750.0F, cont);
                    PointF drawPoinPesT = new PointF(790.0F, cont);
                    //PointF drawadua = new PointF(330.0F, cont + 13);

                    e.Graphics.DrawString(drawLinea, drawFont, drawBrush, drawPointlin);
                    e.Graphics.DrawString(drawProd, drawFont, drawBrush, drawPointProd);
                    e.Graphics.DrawString(drawPedi, drawFont, drawBrush, drawPointPedi);
                    e.Graphics.DrawString(drawcajas, drawFont, drawBrush, drawPointcajas);
                    e.Graphics.DrawString(drawtaraprox, drawFont, drawBrush, drawPointtar);
                    e.Graphics.DrawString(tarima, drawFont3, drawBrush, drawtar);
                    e.Graphics.DrawString(subcli, drawFont3, drawBrush, drawsubcli);
                    string cadTxt = FechaCadTexto(DGV2.Rows[i]);
                    if (cadTxt != "")
                        e.Graphics.DrawString(cadTxt, drawFont3, drawBrush, new PointF(30.0F, cont + 13));
                    //e.Graphics.DrawString(aduan, drawFont, drawBrush, drawadua);
                    //decimal mPeso = TraePesoEst(DGV2.Rows[i].Cells["clave_producto"].Value.ToString());
                    decimal mPeso = peso_promedio(DGV2.Rows[i].Cells["clave_producto"].Value.ToString(), ObtenerProdTipo(DGV2.Rows[i].Cells["clave_producto"].Value.ToString()));
                    e.Graphics.DrawString(mPeso.ToString("#,##0.00"), drawFont3, drawBrush, drawPoinPes);
                    e.Graphics.DrawString((mPeso * Convert.ToInt32(drawcajas)).ToString("#,##0.00"), drawFont3, drawBrush, drawPoinPesT);
                    Totalpeso += mPeso * Convert.ToInt32(drawcajas);
                    cont = cont + 30;

                    PointF drawPoint7 = new PointF(0.0F, cont1);
                    e.Graphics.DrawString(drawString2, drawFont, drawBrush, drawPoint7);
                    cont1 = cont1 + 30;
                    cont2 = cont2 + 30;

                }
                else
                //if (cont_prod == 30)
                //if(i == 24)
                {
                    ultimo_renglon = i;
                    break;
                    //e.HasMorePages = true;
                }
                cont_prod++;
            }//for

            if (cont_prod == 30)
                e.HasMorePages = true;
            else
            {
                e.HasMorePages = false;
                //thisConnection.Open();
                sumapedido = 0; sumataraprox = 0; enteros = 0; sumacomp = 0; sumasplit = 0; decimales = 0;
                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select prod_num_tarimas from tb_cat_producto where prod_clave = '" + Convert.ToString(DGV2.Rows[i].Cells["clave_producto"].Value) + "'";

                    reader2 = cmnd2.ExecuteReader();
                    while (reader2.Read())
                    {
                        prod_num_tarimas = reader2.GetSqlDecimal(0).Value;
                    }
                    reader2.Close();
                    if (prod_num_tarimas == 0)
                    {
                        resulta = 0;
                    }
                    else
                    {
                        resulta = Convert.ToDecimal(DGV2.Rows[i].Cells["cajas"].Value) / prod_num_tarimas;
                    }

                    sumapedido = sumapedido + Convert.ToInt32(DGV2.Rows[i].Cells["cajas"].Value);
                    sumataraprox = sumataraprox + resulta;
                    enteros = Math.Truncate(resulta);
                    sumacomp = sumacomp + enteros;
                    decimales = enteros - resulta;
                    sumasplit = sumasplit + decimales;
                }
                sumatar = sumapedido.ToString();//suma de tarimas
                taraprox = "Tarimas Aprox : " + sumataraprox.ToString("0.00") + "   Total Peso: " + Totalpeso.ToString("#,##0.00");//tarimas aproximadas
                sumcomp = "Completas : " + sumacomp.ToString("0.00") + "    Total Emb " + (Convert.ToDecimal(Totalpeso) + (sumataraprox * 20)).ToString("#,##0.00");//tarimas completas
                sumdeci = "Split : " + (sumasplit * -1).ToString("0.00");//decimales tarimas
                drawobs = "\t" + txtobs.Text;//txt observaciones
                drawPedi = "\t" + txtpedsigma.Text;// txt pedido sigma


                String drawString5 = "\tCargo : __________________________ \t Destino : __________________________ \t Sello : " + txtsello.Text + " \r\n";
                String drawString6 = "\tMe comprometo a mantener la temperatua a : ______________ F \t Nombre del chofer : ______________________ \tFirma Chofer : __________\r\n";
                String drawString7 = "\tTERMOGRAFO : _________________ \t Empezo a cargar (horas) : __________ \t Termino (horas) : __________";
                PointF drawsumatar = new PointF(490.0f, cont1);//suma de tarimas            
                PointF drawaproxtar = new PointF(600.0f, cont1);//tarimas aproximadas
                cont1 = cont1 + 10;
                PointF drawcomptar = new PointF(610.0f, cont1);//tarimas completas
                cont1 = cont1 + 10;
                PointF drawdecitar = new PointF(620.0f, cont1);//tarimas completas
                PointF drawryandig = new PointF(553.0f, 1000.0f);//ryan
                PointF drawobser = new PointF(0.00f, 995.0f);//txt observaciones
                PointF drawsigma = new PointF(0.00f, 1005.0f);//txt pedidosigma
                PointF drawstring5 = new PointF(0.00f, 1010.0f);//string5
                PointF drawchofername = new PointF(505.0f, 1035.0f);//nombre chofer
                PointF drawstring6 = new PointF(0.00f, 1025.0f);//string6
                PointF drawstring7 = new PointF(0.00f, 1040.0f);//string7            
                e.Graphics.DrawString(sumatar, drawFont, drawBrush, drawsumatar);//suma de tarimas
                e.Graphics.DrawString(taraprox, drawFont, drawBrush, drawaproxtar);//tarimas aproximadas
                e.Graphics.DrawString(sumcomp, drawFont, drawBrush, drawcomptar);//tarimas completas
                e.Graphics.DrawString(sumdeci, drawFont, drawBrush, drawdecitar);//tarimas completas
                e.Graphics.DrawString(drawryan, drawFont, drawBrush, drawryandig);//ryan
                e.Graphics.DrawString(drawobs, drawFont, drawBrush, drawobser);//txt observaciones
                e.Graphics.DrawString(drawPedi, drawFont, drawBrush, drawsigma);//txt pedido sigma
                e.Graphics.DrawString(drawString5, drawFont, drawBrush, drawstring5);//string5
                //e.Graphics.DrawString(drawchofer, drawFont, drawBrush, drawchofername);//nombre chofer
                e.Graphics.DrawString(drawString6, drawFont, drawBrush, drawstring6);//string6
                e.Graphics.DrawString(drawString7, drawFont, drawBrush, drawstring7);//string7
                cont1 = 160;
            }
            #endregion
            sumataraprox = sumacomp = sumasplit = decimales = enteros = 0;
            sumapedido = 0;
            thisConnection.Close();
        }

        //boton consulta
        private void btnConsulta_Click(object sender, EventArgs e)
        {
            btnConsulta.Enabled = false;
            btnAlta.Enabled = false;
            btnCancel.Enabled = true;
            txtfolio.Enabled = true;
            txtfolio.ReadOnly = false;
            btnGuardar.Enabled = true;
            btnGuardar.Text = "IMPRIMIR";
            DGV2.ReadOnly = true;
            opcion = 4;
            txtfolio.Focus();
        }

        //trae la información de acuerdo al pedido escrito
        private void txtfolio_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Para obligar a que sólo se introduzcan números 
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = false;
            }
            else
            {
                if (Char.IsControl(e.KeyChar)) //permitir teclas de control como retroceso 
                {
                    e.Handled = false;
                }
                else
                {
                    //el resto de teclas pulsadas se desactivan 
                    e.Handled = true;
                }
            }
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                try
                {

                    //limpia grid de tarimas
                    DataTable DT = (DataTable)DGV2.DataSource;
                    if (DT != null)
                        DT.Clear();
                    //cbdestino.SelectedIndex = -1; 
                    label1.Text = "";
                    thisConnection.Open();
                    cbdestino.Items.Clear();

                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select cnte_clave, pdn_fecha, um_clave, pdn_monto_total, pdn_observacion, pdn_transporte, prov_clave, " +
                                        "pdn_monto_transporte, pdn_pedsigma, pdn_sello, placacaja, pdn_estatus, pdn_tipoflete, pdn_monedaflete, " +
                                        "pdn_pagofletegab, pdn_pedorigen, pdn_ivaflete, pdn_destino, pdn_horent1, cve_subcli, " +
                                        "pdn_tractor, pdn_tractorplaca,pdn_cajatrailer,pdn_transOperador,pdn_cajacapacidad, PDN_ELABORO " +
                                        "from tb_mstr_pedidos_exp " +
                                        "where pdn_folio = '" + txtfolio.Text + "' and pdn_tipo = 'EXP'";
                    reader2 = cmnd2.ExecuteReader();
                    while (reader2.Read())
                    {
                        txtcliente.Text = reader2.GetSqlString(0).ToString().Trim();
                        txtfecha.Text = reader2.GetSqlDateTime(1).ToString().Substring(0, 10);
                        //txtfecha.Text = txtfecha.Text.Substring(0,10);
                        um_clave = reader2.GetSqlString(2).ToString().Trim();
                        if (um_clave.Trim() == "PESOS")
                            txtmn.Text = Convert.ToString(reader2.GetSqlDecimal(3).Value).Trim();

                        if (um_clave.Trim() == "USD")
                            txtusd.Text = Convert.ToString(reader2.GetSqlDecimal(3).Value).Trim();

                        if (um_clave.Trim() == "EUROS")
                            txteuro.Text = Convert.ToString(reader2.GetSqlDecimal(3).Value).Trim();

                        txtobs.Text = reader2.GetSqlString(4).ToString().Trim();
                        txttalonemb.Text = reader2.GetSqlString(5).ToString().Trim();
                        txtclavetrans.Text = reader2.GetSqlString(6).ToString().Trim();
                        txtimporte.Text = Convert.ToString(reader2.GetSqlDecimal(7).Value).Trim();
                        txtpedsigma.Text = reader2.GetSqlString(8).ToString().Trim();
                        txtsello.Text = reader2.GetSqlString(9).ToString().Trim();
                        txtcajatrailer.Text = reader2.GetSqlString(10).ToString().Trim();
                        pdn_estatus = reader2.GetSqlString(11).ToString().Trim();
                        CBtipoflete.SelectedItem = reader2.GetValue(12).ToString().Trim();
                        tipoflete = reader2.GetValue(12).ToString().Trim();
                        CBmonfle.SelectedItem = reader2.GetValue(13).ToString().Trim();
                        monedaflete = reader2.GetValue(13).ToString().Trim();
                        flete_gab = reader2.GetValue(14).ToString().Trim();
                        if (flete_gab.Contains("S"))
                            cbpagoGAB.Checked = true;
                        txtcompleped.Text = reader2.GetValue(15).ToString().Trim();
                        iva_flete = reader2.GetValue(16).ToString().Trim();
                        if (reader2.GetValue(16).ToString().Trim().Contains("S"))
                        {
                            cbIVA.Checked = true;
                            //txtsubtotal.Text = (Convert.ToDecimal(txtimporte.Text) / Convert.ToDecimal(1.12)).ToString("###,###,##0.00");
                            //txtIVAfle.Text = (Convert.ToDecimal(txtsubtotal.Text) / Convert.ToDecimal(0.16)).ToString("###,###,##0.00");
                            //txtretension.Text = (Convert.ToDecimal(txtsubtotal.Text) / Convert.ToDecimal(0.04)).ToString("###,###,##0.00");
                        }
                        //int index = cbdestino.FindStringExact(reader2.GetValue(17).ToString().Trim()); esto estaba antes 9nov2018 rcc
                        //int index = cbdestino.FindStringExact(reader2.GetValue(19).ToString().Trim());
                        //cbdestino.SelectedIndex = index;
                        string query = "select cnte_clave, entr_clave, entr_nombre, Rtrim(entr_NOMBRE) + ' ' + RTRIM(ENTR_CIUDAD) AS NOMBRE  from tb_cat_entregas order by entr_nombre";
                        da = new SqlDataAdapter(query, thisConnection);
                        da.Fill(ds1, "sub");
                        subcnte = ds1.Tables["sub"];
                        CveSubCli = reader2.GetValue(19).ToString().Trim();
                        cbdestino.Text = "";
                        foreach (DataRow dr in subcnte.Select("cnte_clave = '" + txtcliente.Text + "' and entr_clave = '" + CveSubCli + "'"))
                        {
                            cbdestino.Items.Add(dr["entr_nombre"].ToString().Trim());
                            cbdestino.SelectedIndex = 0;
                        }
                        string HrCita = reader2["pdn_horent1"].ToString();
                        DtHrCita.Text = (HrCita.Trim().Length == 0) ? "00:00" : HrCita;
                        TxtTractor.Text = reader2["pdn_tractor"].ToString().Trim();
                        TxtCjTracto.Text = reader2["pdn_tractorplaca"].ToString().Trim();
                        TxtCjTra.Text = reader2["pdn_cajatrailer"].ToString().Trim();
                        TxtChofer.Text = reader2["pdn_transOperador"].ToString().Trim();
                        TxtCapaCaja.Text = reader2["pdn_cajacapacidad"].ToString().Trim();
                        LblElaboro.Text = "";
                        foreach (DataRow row in CatUser.Select("usu_login = '" + reader2["pdn_elaboro"].ToString().Trim() + "'"))
                            LblElaboro.Text = "Elaboro: " + row["usu_nombre"].ToString().Trim();
                    }
                    if (reader2.HasRows == false)
                    {
                        thisConnection.Close();
                        MessageBox.Show("La orden de venta de exportación no existe", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    reader2.Close();

                    foreach (DataRow dr in subcnte.Select("cnte_clave = '" + txtcliente.Text + "'"))
                    {
                        subcliente.Items.Add(dr["nombre"].ToString().Trim());//entr_nombre cambio 28062023
                    }

                    cmnd2 = thisConnection.CreateCommand();
                    if (um_clave.Trim() == "PESOS")
                        cmnd2.CommandText = "SELECT tdp.lin_clave as linea, tdp.prod_clave as producto, tcpr.prod_nombre AS nombre, tdp.pdn_num_unidades as cantidad, " +
                                            "tdp.pdn_precio_mn as precio, tdp.pdn_tipotar as tipo_tar, tdp.pdn_subcli as subcliente, tdp.pdn_ordventa as po_number, " +
                                            "tdp.pdn_observaciones as observaciones, (tdp.pdn_num_unidades * tdp.pdn_precio_mn) as total, tdp.pdn_descrip_especial as special, " +
                                            "tdp.pdn_num_ped as pedimento,tcpr.prod_fraccion,tcpr.prod_peso, tdp.pdn_fecha_cad FROM tb_det_pedidos tdp, tb_cat_producto tcpr " +
                                            "WHERE tdp.pdn_folio = '" + txtfolio.Text + "' AND tcpr.prod_clave = tdp.prod_clave and tdp.pdn_tipo= 'EXP' ORDER BY tdp.nrenglon ASC";

                    if (um_clave.Trim() == "USD")
                        cmnd2.CommandText = "SELECT tdp.lin_clave as linea, tdp.prod_clave as producto, tcpr.prod_nombre AS nombre, tdp.pdn_num_unidades as cantidad, " +
                                             "tdp.pdn_precio_usd as precio, tdp.pdn_tipotar as tipo_tar, tdp.pdn_subcli as subcliente, tdp.pdn_ordventa as po_number, " +
                                             "tdp.pdn_observaciones as observaciones, (tdp.pdn_num_unidades * tdp.pdn_precio_usd) as total, tdp.pdn_descrip_especial as special, " +
                                             "tdp.pdn_num_ped as pedimento,tcpr.prod_fraccion,tcpr.prod_peso, tdp.pdn_fecha_cad FROM tb_det_pedidos tdp, tb_cat_producto tcpr " +
                                             "WHERE tdp.pdn_folio = '" + txtfolio.Text + "' AND tcpr.prod_clave = tdp.prod_clave and tdp.pdn_tipo= 'EXP' ORDER BY tdp.nrenglon ASC";

                    if (um_clave.Trim() == "EURO" || um_clave.Trim() == "EUROS")
                        cmnd2.CommandText = "SELECT tdp.lin_clave as linea, tdp.prod_clave as producto, tcpr.prod_nombre AS nombre, tdp.pdn_num_unidades as cantidad, " +
                                             "tdp.pdn_precio_usd as precio, tdp.pdn_tipotar as tipo_tar, tdp.pdn_subcli as subcliente, tdp.pdn_ordventa as po_number, " +
                                             "tdp.pdn_observaciones as observaciones, (tdp.pdn_num_unidades * tdp.pdn_precio_usd) as total, tdp.pdn_descrip_especial, " +
                                             "tdp.pdn_num_ped as pedimento,tcpr.prod_fraccion,tcpr.prod_peso, tdp.pdn_fecha_cad FROM tb_det_pedidos tdp, tb_cat_producto tcpr " +
                                             "WHERE tdp.pdn_folio = '" + txtfolio.Text + "' AND tcpr.prod_clave = tdp.prod_clave and tdp.pdn_tipo= 'EXP' ORDER BY tdp.nrenglon ASC";

                    reader2 = cmnd2.ExecuteReader();
                    habilita(false);
                    Fraccion.Rows.Clear();
                    while (reader2.Read())
                    {
                        DataRow drc = tmgrdped.NewRow();
                        drc["linea"] = reader2.GetSqlString(0).ToString().Trim();
                        drc["producto"] = reader2.GetSqlString(1).ToString().Trim();
                        drc["nombre"] = reader2.GetSqlString(2).ToString().Trim();
                        drc["cantidad"] = reader2.GetSqlDecimal(3).Value.ToString("0");
                        drc["precio"] = reader2.GetSqlDecimal(4).ToString();
                        drc["clave_tarima"] = reader2.GetSqlString(5).ToString().Trim();
                        drc["subcnte_clave"] = reader2.GetSqlString(6).Value.Trim();
                        drc["po_number"] = reader2.GetSqlString(7).ToString().Trim();
                        drc["observaciones"] = reader2.GetSqlString(8).ToString().Trim();
                        drc["total"] = reader2.GetSqlDecimal(9).Value.ToString("###,###,###.00");
                        foreach (DataRow dr in subcnte.Select("cnte_clave = '" + txtcliente.Text + "' and entr_clave = '" + reader2.GetValue(6).ToString().Trim() + "'"))
                        {
                            drc["subcliente"] = Convert.ToString(dr["entr_nombre"].ToString().Trim());
                        }
                        foreach (DataRow dr in tarimas.Select("emp_clave = '" + reader2.GetValue(5).ToString().Trim() + "'"))
                        {
                            drc["nom_tarima"] = Convert.ToString(dr["emp_nombre"].ToString().Trim());
                        }
                        drc["transito"] = "";
                        drc["especial"] = "Obs Especial";
                        drc["special"] = reader2.GetValue(10).ToString();
                        drc["pedimento"] = reader2.GetValue(11).ToString();
                        drc["fecha_cad"] = reader2.IsDBNull(14) ? "" : Convert.ToDateTime(reader2.GetValue(14)).ToString("dd/MM/yyyy");
                        tmgrdped.Rows.Add(drc);
                        string Aduana = "01";
                        if (reader2.GetSqlString(1).ToString().Trim() == "17TERMOGRA")
                        {
                            Aduana = "06";
                        }
                        decimal Cant = Convert.ToDecimal(reader2.GetSqlDecimal(3).Value.ToString());
                        decimal Peso = (reader2.GetSqlDecimal(13).Value.ToString().Trim() == "") ? 0 : Convert.ToDecimal(reader2.GetSqlDecimal(13).Value.ToString().Trim());
                        decimal Kilos = Peso * Cant;
                        decimal Imp = Convert.ToDecimal(reader2.GetSqlDecimal(9).Value.ToString());
                        decimal PreUni = 0, Imp2 = 0;
                        if (Kilos != 0)
                        {
                            PreUni = Math.Round(Convert.ToDecimal(Imp / Kilos), 4);
                            //decimal PreUni = Math.Round(Convert.ToDecimal(Kilos / Imp), 2);
                            //Imp2 = Imp / PreUni;
                            Imp2 = Imp / PreUni;
                        }
                        Fraccion.Rows.Add(reader2.GetSqlString(1).ToString().Trim(), reader2.GetSqlString(2).ToString().Trim(), reader2.GetSqlString(12).ToString().Trim(), Aduana, Cant, Peso, Imp2, PreUni, Imp, Kilos);
                    }
                    reader2.Close();
                    DGV2.DataSource = tmgrdped;
                    thisConnection.Close();
                    txtcliente_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                    txtclavetrans_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));

                    //identifica el tipo de moneda
                    if (um_clave.Trim() == "" || um_clave.Trim() == "PESOS")
                    {
                        CBmoneda.SelectedIndex = 0;
                        DGV2.Columns["precios"].HeaderText = "Precio M.N.";//4
                        DGV2.Columns["total"].HeaderText = "Total M.N.";//10 
                        moneda = "PESOS";
                    }
                    if (um_clave.Trim() == "USD")
                    {
                        CBmoneda.SelectedIndex = 1;
                        DGV2.Columns["precios"].HeaderText = "Precio USD";//4
                        DGV2.Columns["total"].HeaderText = "Total USD";//10
                        //DGV2.Columns[11].Visible = false;                        
                        moneda = "USD";
                    }
                    if (um_clave.Trim() == "EURO")
                    {
                        CBmoneda.SelectedIndex = 1;
                        DGV2.Columns["precios"].HeaderText = "Precio EURO";//4
                        DGV2.Columns["total"].HeaderText = "Total EURO";//10
                        //DGV2.Columns[11].Visible = false;                        
                        moneda = "EURO";
                    }
                    //CBmoneda_SelectionChangeCommitted(sender, e);
                    CBmoneda.SelectedItem = um_clave.Trim();
                    txtclaveprod.Enabled = false;
                    CBProducto.Enabled = false;
                    //DGV2.Columns[12].Visible = false;
                    btnActu.Enabled = true;
                    btnbaja.Enabled = true;

                    for (int i = 0; i < DGV2.Rows.Count; i++)
                    {
                        this.DGV2.Rows[i].HeaderCell.Value = (i + 1).ToString();
                    }
                    DGV2.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader);

                    if (pdn_estatus.Trim() == "C")
                    {
                        btnbaja.Enabled = false;
                        btnActu.Enabled = false;
                        MessageBox.Show("Orden de exportanción cancelada", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        label1.Text = "Estatus: CANCELADA";
                    }
                    if (pdn_estatus.Trim() == "F")
                    {
                        btnbaja.Enabled = false;
                        //btnActu.Enabled = false;
                        DGV2.ReadOnly = true;
                        MessageBox.Show("Orden de exportanción facturada", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        label1.Text = "Estatus: FACTURADA";
                    }
                    //txtcliente_KeyPress(sender, e);
                    for (int i = 0; i < DGV2.Rows.Count; i++)
                    {
                        string txtnumped = Convert.ToString(DGV2.Rows[i].Cells["pedimento"].Value).Trim();
                        if ((txtnumped.Trim() != "") || (txtnumped.Contains(' ')))
                        {
                            for (int j = 0; j < Convert.ToString(DGV2.Rows[i].Cells["pedimento"].Value).Trim().Length; j++)
                            {
                                string pedimento = txtnumped;
                                if (j == 2)
                                    txtnumped = pedimento.Insert(2, "  ");
                                if (j == 6)
                                    txtnumped = pedimento.Insert(6, "  ");
                                if (j == 12)
                                    txtnumped = pedimento.Insert(12, "  ");
                            }
                            DGV2.Rows[i].Cells["pedimento"].Value = txtnumped;
                        }
                    }
                }
                catch (SqlException ex)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                    MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                catch (Exception ex1)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                    MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }

        //botón de actualizar
        private void btnActu_Click(object sender, EventArgs e)
        {
            string estadoorden = ValidarOrdenVentaEmbarque();
            if (estadoorden == "Iniciado")
            {
                Utilerias.Class1.registrar_movimiento3(DateTime.Now,
                                        Environment.MachineName,
                                        Utilerias.Class1.Usu_login,
                                        "C",
                                        "4.1",
                                        txtfolio.Text,
                                        "La Orden de venta EXP tiene embarque Iniciado, Considere avisar a Embarques de Cualquier cambio y considere que los tiempos de Carga seran afectados",
                                        "SIPGAB",
                                        txtfolio.Text,
                                        "Se muestra el aviso al usuario");
                MessageBox.Show("La Orden de venta tiene embarque Iniciado, Considere avisar a Embarques de Cualquier cambio y considere que los tiempos de Carga seran afectados", "Carga Iniciada", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (estadoorden == "Finalizado")
            {
                MessageBox.Show("La Orden de venta tiene embarque Terminado y cerrado. No se pueden realizar Modificaciones hasta que la orden sea abierta", "Carga Terminada", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }




            thisConnection.Open();
            string hora_fin = "", sts = "";
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "select hora_fin, sts from tb_mstr_embarque where emb_folio = '" + txtfolio.Text + "' and emb_tipo = 'EXP'";
            reader2 = cmnd2.ExecuteReader();
            while (reader2.Read())
            {
                hora_fin = reader2.GetValue(0).ToString().Trim();
                sts = reader2.GetValue(1).ToString().Trim();
            }
            reader2.Dispose();
            thisConnection.Close();

            if ((hora_fin.Trim() != "--:--") && (sts.Trim() == "T"))
            {
                MessageBox.Show("El embarque ya esta cerrado. Cualquier modificación favor de avisar a embarques", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DGV2.Columns["cajas"].ReadOnly = true;
            }

            btnbaja.Enabled = false;
            btnActu.Enabled = false;
            btnGuardar.Enabled = true;
            habilita(true);
            CBCliente.Enabled = false;
            txtcliente.Enabled = false;
            CBProducto.Enabled = true;
            txtclaveprod.Enabled = true;
            txtfecha.Enabled = true;
            thisConnection.Open();
            transportista();
            thisConnection.Close();
            Faltantes.Clear();
            opcion = 3;
            btnGuardar.Text = "GUARDAR";
            DGV2.ReadOnly = false;
            DGV2.Columns[0].ReadOnly = true;
            DGV2.Columns[1].ReadOnly = true;
            DGV2.Columns[2].ReadOnly = true;
            txtcliente.Enabled = true;
            CBCliente.Enabled = true;

            if (pdn_estatus.Trim() == "F")
            {
                DGV2.ReadOnly = true;
            }
        }

        //boton baja
        private void btnbaja_Click(object sender, EventArgs e)
        {
            btnbaja.Enabled = false;
            btnActu.Enabled = false;
            btnGuardar.Enabled = true;
            //DGV2.ReadOnly = false;
            opcion = 2;
            btnGuardar.Text = "GUARDAR";
            //printDialog1.Document = new PrintDocument();
            //PrintPreviewDialog VistaPrevia = new PrintPreviewDialog();
            //printDialog1.Document.PrintPage += printDocument1_PrintPage;
            //VistaPrevia.Document = printDialog1.Document;
            //VistaPrevia.ShowDialog();
        }

        //botón cancelar
        private void btnCancel_Click(object sender, EventArgs e)
        {
            DGV2.ReadOnly = false;
            limpiarTextBoxes(this);
            habilita(false);
            btnbaja.Enabled = false;
            btnActu.Enabled = false;
            btnGuardar.Text = "GUARDAR";
            btncopiarpedido.Enabled = false;
            BtnPedTom.Enabled = false;
            cbIVA.Checked = false;
            cbpagoGAB.Checked = false;
            CBmonfle.SelectedIndex = -1;
            CBtipoflete.SelectedIndex = -1;
            lin_clave = ""; prod_clave = ""; prod_nombre = ""; cnte_sucursal = ""; cnte_entregas = ""; pdn_folio = ""; ArtCod = " "; ArtNom = ""; cveunidad = " ";
            sig = ""; id_cliente = ""; id_vendedor = ""; um_clave = ""; ultimofolio = ""; cve_subcli = ""; moneda = ""; pdn_folio_web = ""; vend_clave = ""; vend_nombre = "";
            ryandigital = " "; ryantipo = ""; pdn_pedimento = ""; pdn_fechap = ""; pdn_aduana = ""; pdn_pre_cad = ""; pdn_cantxcaja = ""; pdn_tipo = ""; pdn_estatus = "";
            flete_gab = "N"; monedaflete = ""; tipoflete = ""; iva_flete = "N";
            precio = 0; preciouni = 0; resulta = 0; prod_num_tarimas = 0; sumataraprox = 0; sumacomp = 0; sumasplit = 0; enteros = 0; decimales = 0;
            prod_precio_mn = 0; prod_precio_usd = 0; var_peso_unitario = 0; var_dec_peso_total = 0; pdn_num_unidades = 0; pdn_precio_mn = 0;
            pdn_precio_usd = 0; moneyusd = 0; moneymn = 0; moneyeuro = 0;
            encontro = false;
            sige = 0; cantidad = 0; sumapedido = 0; opcion = 0; pedorigen = 0;
            cont_prod = 1; ultimo_renglon = 0;
            lbpedido.Visible = false;
            txtpedido.Visible = false;
        }

        private void PBbtnsalir_Click(object sender, EventArgs e)
        {
            //printDialog1.Document = new PrintDocument();
            //PrintPreviewDialog VistaPrevia = new PrintPreviewDialog();
            //printDialog1.Document.PrintPage += printDocument1_PrintPage;
            //VistaPrevia.Document = printDialog1.Document;
            //VistaPrevia.ShowDialog();
            try
            {
                thisConnection.Open();
                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "SELECT TOP 1 inicio_sesion, usu_login FROM tb_cat_historial_dia where nombre_maquina = '" + Environment.MachineName + "' and sistema= 'SIPGAB' ORDER BY inicio_sesion desc";
                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    Utilerias.Class1.Inicio_sesion = reader2.GetSqlDateTime(0).Value;
                    Utilerias.Class1.Usu_login = reader2.GetSqlString(1).ToString();
                    Utilerias.Class1.Nombre_equipo = Environment.MachineName;
                }
                reader2.Close();

                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "update tb_cat_historial_dia set formulario = ' ' where nombre_maquina ='" + Utilerias.Class1.Nombre_equipo + "' and usu_login = '" + Utilerias.Class1.Usu_login + "' and inicio_sesion = '" + Utilerias.Class1.Inicio_sesion.ToString("s") + "' and sistema ='SIPGAB'";
                reader2 = cmnd2.ExecuteReader();
                reader2.Close();
                thisConnection.Close();
                Application.Exit();
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex1)
            {
                thisConnection.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void btnsubcliente_Click(object sender, EventArgs e)
        {
            Alta_Subclientes AS = new Alta_Subclientes();
            AS.ShowDialog();
        }

        private void DGV2_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception.Message == "El valor de DataGridViewComboBoxCell no es válido." || e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = DGV2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)DGV2.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    ((DataGridViewComboBoxColumn)DGV2.Columns[e.ColumnIndex]).Items.Add(value);
                    e.ThrowException = false;
                }
            }
        }


        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DGV2.CurrentCell.ColumnIndex == 9) // combo subcliente
            {
                string nombre = cntesub.SelectedItem.ToString().Trim();
                foreach (DataRow dr in subcnte.Select("nombre = '" + nombre + "' and cnte_clave = '" + txtcliente.Text + "'"))
                {
                    DGV2.Rows[DGV2.CurrentRow.Index].Cells["subcnte_clave"].Value = Convert.ToString(dr["entr_clave"].ToString().Trim());
                    DGV2.Rows[DGV2.CurrentRow.Index].Cells[9].Value = nombre;
                    DGV2.Rows[DGV2.CurrentRow.Index].Cells[10].Selected = true;
                    return;
                }
                DGV2.Rows[DGV2.CurrentRow.Index].Cells[9].Value = nombre;
                DGV2.Rows[DGV2.CurrentRow.Index].Cells[10].Selected = true;
                return;
            }
            if (DGV2.CurrentCell.ColumnIndex == 7)
            {
                ComboBox combo = sender as ComboBox;
                string nombre = combo.SelectedItem.ToString().Trim();
                DataGridViewRow row = DGV2.CurrentRow;
                foreach (DataRow dr in tarimas.Select("emp_nombre = '" + nombre + "'"))
                {
                    tme = Convert.ToString(dr["emp_clave"].ToString().Trim());
                }
                //string nombre = tarim.SelectedItem.ToString().Trim();
                //DGV2.Rows[DGV2.CurrentRow.Index].Cells["tipo_tar"].Value = tme;
                //foreach (DataRow dr in tarimas.Select("emp_nombre = '" + nombre + "'"))
                //{ 
                //    DGV2.Rows[DGV2.CurrentRow.Index].Cells["tipo_tar"].Value = Convert.ToString(dr["emp_clave"].ToString()).Trim();
                //}
            }
        }

        private void btncajatrailer_Click(object sender, EventArgs e)
        {
            CajaTrailer ct = new CajaTrailer();
            ct.ShowDialog();
            txtcajatrailer.Text = Utilerias.Class1.Nombre_equipo;
        }

        private void txtimporte_TextChanged(object sender, EventArgs e)
        {
            if (txtimporte.Text.Trim() == "")
                txtimporte.Text = "0.00";
            //else
            //{
            //    if (Convert.ToDecimal(txtimporte.Text) != 0)
            //    {
            //        txtsubtotal.Text = (Convert.ToDecimal(txtimporte.Text) / Convert.ToDecimal(1.12)).ToString("###,###,##0.00");
            //    }
            //}
        }

        private void btnminimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void cbIVA_CheckedChanged(object sender, EventArgs e)
        {
            if (cbIVA.Checked == true)
            {
                if (txtimporte.Text.Trim() != "")
                {
                    txtsubtotal.Text = (Convert.ToDecimal(txtimporte.Text) / Convert.ToDecimal(1.12)).ToString("###,###,##0.00");
                    txtIVAfle.Text = (Convert.ToDecimal(txtsubtotal.Text) * Convert.ToDecimal(0.16)).ToString("###,###,##0.00");
                    txtretension.Text = (Convert.ToDecimal(txtsubtotal.Text) * Convert.ToDecimal(0.04)).ToString("###,###,##0.00");
                }
                else
                {
                    MessageBox.Show("Favor de escribir el importe del flete", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtimporte.Focus();
                    cbIVA.Checked = false;
                    return;
                }
            }

            else
            {
                txtsubtotal.Text = "";
                txtIVAfle.Text = "";
                txtretension.Text = "";
            }
        }

        private string ValidaPedidoOrigen()
        {
            string Existe = "N";
            thisConnection.Open();
            Int32 MFol = Convert.ToInt32(txtcompleped.Text);
            string BD = "tb_mstr_pedidos_nal";
            if (MFol < 300000)
                BD = "tb_mstr_pedidos_exp";
            string Cadena = "SELECT pdn_pedorigen, pdn_fecha from " + BD + " WHERE pdn_folio = '" + MFol + "'";
            SqlCommand cmd = new SqlCommand(Cadena, thisConnection);
            reader2 = cmd.ExecuteReader();
            string pedidoorigen = "";
            DateTime FecPed = System.DateTime.Today;
            string Encontro = "F";
            while (reader2.Read())
            {
                pedidoorigen = Convert.ToString(reader2["pdn_pedorigen"]).Trim();
                FecPed = Convert.ToDateTime(reader2["pdn_fecha"]);
                Encontro = "S";
            }
            thisConnection.Close();
            if (pedidoorigen.Trim() != "0" && pedidoorigen.Trim() != "")
                Existe = pedidoorigen.Trim();
            if (Encontro == "S")
                if (FecPed != Convert.ToDateTime(txtfecha.Text))
                {
                    MessageBox.Show("Error: EL PEDIDO ORIGEN TIENE DIFERENTE FECHA AL PEDIDOS COMPLEMENTO " + txtcompleped.Text + System.Environment.NewLine + " FECHA DEL PEDIDO COMPLEMENTO:  " + FecPed, "ERROR EN LAS FECHAS DE LOS PEDIDOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Existe = "N";
                }
            return Existe;
        }

        private void txtcompleped_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (txtcompleped.Text != "")
                {
                    string estadoorden = ValidarOrdenVentaEmbarqueComplemento();
                    if (estadoorden == "Iniciado")
                    {
                        MessageBox.Show("La Orden de Venta Origen tiene embarque Iniciado, Considere avisar a Embarques de Cualquier cambio y considere que los tiempos de Carga seran afectados", "Carga Iniciada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (estadoorden == "Finalizado")
                    {
                        MessageBox.Show("La Orden de Venta Origen tiene embarque Terminado y cerrado. No se pueden realizar Modificaciones a los productos hasta que la orden sea abierta", "Carga Terminada", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        txtcompleped.Focus();
                        txtcompleped.SelectAll();
                        return;
                    }
                }

                string pedidoorigen = "N";

                string pedidoCargado = "N";

                pedidoCargado = ValidaCarga();

                if (pedidoCargado != "N")
                {
                    if (pedidoCargado == "C")
                    {
                        MessageBox.Show("Error: EL PEDIDO ACTUAL YA ESTA EN INICIO DE CARGA, FAVOR DE AVISAR A EMBARQUES PARA CANCELARLO, DESCARGAR EL PEDIDO Y REALICE LA MODIFICACION NUEVAMENTE", "PEDIDO EN PROCESO DE CARGA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //txtcompleped.Text = pedidocomplemento;
                        return;
                    }
                    else if (pedidoCargado == "T")
                    {
                        MessageBox.Show("Error: ESTE PEDIDO YA SE HA CERRADO EN LA CARGA, NO SE PUEDE CAMBIAR EL PEDIDO ORIGEN/COMPLEMENTO", "PEDIDO YA CARGADO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //txtcompleped.Text = pedidocomplemento;
                        return;
                    }
                }

                if (txtcompleped.Text.Length > 0 && txtcompleped.Text != "0")
                {
                    pedidoorigen = ValidaPedidoOrigen();
                }
                if (pedidoorigen != "N")
                {
                    MessageBox.Show("Error: ESTE PEDIDO " + txtcompleped.Text + System.Environment.NewLine + "NO ES EL PEDIDO ORIGEN/PRINCIPAL, EL PEDIDO ESTA ASOCIADO A LA ORDEN: " + pedidoorigen, "PEDIDO NO ES ORIGEN/PRINCIPAL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtcompleped.Text = "";
                    return;
                }

                if (txtcompleped.Text.Trim() == txtfolio.Text.Trim())
                {
                    MessageBox.Show("Error: ESTE PEDIDO " + txtcompleped.Text + System.Environment.NewLine + " Es Igual Al Pedido Actual, El pedido Principal/Origen no puede ser el mismo que el pedido actual", "PEDIDO IGUAL AL PEDIDO PRINCIPAL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtcompleped.Text = "";
                    return;
                }

                thisConnection.Open();
                cmnd2 = thisConnection.CreateCommand();
                if (Convert.ToInt32(txtcompleped.Text) > 300000)
                    cmnd2.CommandText = "select a.prov_clave, a.placacaja, a.pdn_transporte, a.pdn_sello, b.prov_nombre, a.pdn_monedaflete, " +
                                        "a.pdn_tipoflete, a.pdn_ivaflete, a.pdn_pagofletegab, a.pdn_horent1 from " +
                                        "tb_mstr_pedidos_nal a, tb_cat_proveedor b " +
                                        "where a.pdn_folio = '" + txtcompleped.Text + "' and b.prov_clave = a.prov_clave";

                else
                    cmnd2.CommandText = "select a.prov_clave, a.placacaja, a.pdn_transporte, a.pdn_sello, b.prov_nombre, a.pdn_monedaflete, " +
                                    "a.pdn_tipoflete, a.pdn_ivaflete, a.pdn_pagofletegab,a.pdn_tractor,a.pdn_tractorplaca, a.pdn_cajatrailer,a.pdn_transoperador,a.pdn_cajacapacidad, a.pdn_horent1 from " +
                                    "tb_mstr_pedidos_exp a, tb_cat_proveedor b " +
                                    "where a.pdn_folio = '" + txtcompleped.Text + "' and b.prov_clave = a.prov_clave";

                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    txtclavetrans.Text = reader2.GetValue(0).ToString().Trim();
                    txtcajatrailer.Text = reader2.GetValue(1).ToString().Trim();
                    txttalonemb.Text = reader2.GetValue(2).ToString().Trim();
                    txtsello.Text = reader2.GetValue(3).ToString().Trim();
                    CBTransportista.SelectedItem = reader2.GetValue(4).ToString();
                    CBmonfle.SelectedItem = reader2.GetValue(5).ToString().Trim();
                    CBtipoflete.SelectedItem = reader2.GetValue(6).ToString().Trim();
                    try
                    {
                        TxtTractor.Text = reader2["pdn_tractor"].ToString();
                        TxtCjTracto.Text = reader2["pdn_tractorplaca"].ToString();
                        TxtCjTra.Text = reader2["pdn_cajatrailer"].ToString();
                        TxtChofer.Text = reader2["pdn_transoperador"].ToString();
                        TxtCapaCaja.Text = reader2["pdn_cajacapacidad"].ToString();
                    }
                    catch
                    {

                    }
                    DtHrCita.Value = Convert.ToDateTime(reader2["pdn_horent1"].ToString());
                    if (reader2.GetValue(7).ToString().Trim().Contains("S"))
                        cbIVA.Checked = true;
                    else
                        cbIVA.Checked = false;
                    if (reader2.GetValue(8).ToString().Trim().Contains("S"))
                        cbpagoGAB.Checked = true;
                    else
                        cbpagoGAB.Checked = false;
                }
                reader2.Close();
                thisConnection.Close();
            }
        }

        private string ValidaCarga()
        {
            string Existe = "N";
            thisConnection.Open();
            Int32 MFol = Convert.ToInt32(txtfolio.Text.Trim());
            string BD = "tb_mstr_pedidos_nal";
            if (MFol < 300000)
                BD = "tb_mstr_pedidos_exp";
            string Cadena = "SELECT pdn_pedorigen from " + BD + " WHERE pdn_folio = '" + MFol + "'";
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnection;
            string pedidoorigen = Convert.ToString(cmd.ExecuteScalar()).Trim();
            thisConnection.Close();
            if (pedidoorigen.Trim() != txtcompleped.Text.Trim())
            {
                thisConnection.Open();
                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "select sts FROM tb_mstr_embarque WHERE emb_folio = '" + txtfolio.Text.Trim() + "'";
                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    Existe = reader2.GetValue(0).ToString().Trim();
                }
                thisConnection.Close();
            }
            return Existe;
        }

        private void CBProducto_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.ToString() == e.KeyChar.ToString().ToUpper())
            {
                return;
            }
            e.Handled = true;
            SendKeys.Send(e.KeyChar.ToString().ToUpper());
        }

        private void CBCliente_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.ToString() == e.KeyChar.ToString().ToUpper())
            {
                return;
            }
            e.Handled = true;
            SendKeys.Send(e.KeyChar.ToString().ToUpper());
        }

        private void CBmoneda_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.ToString() == e.KeyChar.ToString().ToUpper())
            {
                return;
            }
            e.Handled = true;
            SendKeys.Send(e.KeyChar.ToString().ToUpper());
        }

        private void CBTransportista_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.ToString() == e.KeyChar.ToString().ToUpper())
            {
                return;
            }
            e.Handled = true;
            SendKeys.Send(e.KeyChar.ToString().ToUpper());
        }

        private void CBtipoflete_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.ToString() == e.KeyChar.ToString().ToUpper())
            {
                return;
            }
            e.Handled = true;
            SendKeys.Send(e.KeyChar.ToString().ToUpper());
        }

        private void CBmonfle_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.ToString() == e.KeyChar.ToString().ToUpper())
            {
                return;
            }
            e.Handled = true;
            SendKeys.Send(e.KeyChar.ToString().ToUpper());
        }

        private void DGV2_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void CBmoneda_KeyDown(object sender, KeyEventArgs e)
        {
            CBmoneda.DroppedDown = true;
        }

        private void CBCliente_KeyDown(object sender, KeyEventArgs e)
        {
            CBCliente.DroppedDown = true;
        }

        private void CBProducto_KeyDown(object sender, KeyEventArgs e)
        {
            CBProducto.DroppedDown = true;
        }

        private void CBTransportista_KeyDown(object sender, KeyEventArgs e)
        {
            CBTransportista.DroppedDown = true;
        }

        private void CBmonfle_KeyDown(object sender, KeyEventArgs e)
        {
            CBtipoflete.DroppedDown = true;
        }

        private void CBtipoflete_KeyDown(object sender, KeyEventArgs e)
        {
            CBtipoflete.DroppedDown = true;
        }

        private void txtcompleped_TextChanged(object sender, EventArgs e)
        {
            if (txtcompleped.Text.Trim() == "")
                txtcompleped.Text = "0";
        }

        private void CreaQr(string Folio)
        {
            BarcodeLib.Barcode.QRCode qrbarcode = new BarcodeLib.Barcode.QRCode();

            // Select QR Code data encoding type: numeric, alphanumeric, byte, and Kanji to select from.
            qrbarcode.Encoding = BarcodeLib.Barcode.QRCodeEncoding.Auto;
            qrbarcode.Data = Folio;

            // Adjusting QR Code barcode module size and quiet zones on four sides.
            qrbarcode.ModuleSize = 1;
            qrbarcode.LeftMargin = 1; //12
            qrbarcode.RightMargin = 1; //12 
            qrbarcode.TopMargin = 1; //12
            qrbarcode.BottomMargin = 1; //12

            // Select QR Code Version (Symbol Size), available from V1 to V40, i.e. 21 x 21 to 177 x 177 modules.
            qrbarcode.Version = BarcodeLib.Barcode.QRCodeVersion.V1;

            // Set QR-Code bar code Reed Solomon Error Correction Level: L(7%), M (15%), Q(25%), H(30%)
            qrbarcode.ECL = BarcodeLib.Barcode.QRCodeErrorCorrectionLevel.H; //L
            qrbarcode.ImageFormat = System.Drawing.Imaging.ImageFormat.Png;

            // More barcode settings here, like ECI, FNC1, Structure Append, etc.

            // save barcode image into your system
            qrbarcode.drawBarcode(@"c:/reportes/qrcode.png");

            // Generate QR Code barcode & output to byte array
            //byte[] barcodeInBytes = qrbarcode.drawBarcodeAsBytes();

            // Generate QR Code barcode to Graphics object
            //Graphics graphics = ...;
            //qrbarcode.drawBarcode(graphics);

            // Generate QR Code barcode and output to HttpResponse object
            //HttpResponse response = ...;
            //qrbarcode.drawBarcode(response);

            // Generate QR Code barcode and output to Stream object
            //Stream stream = "@c:\reporte\qr";
            //qrbarcode.drawBarcode(stream);
        }

        private void txtcompleped_Leave(object sender, EventArgs e)
        {
            if (txtcompleped.Text != "")
            {
                string estadoorden = ValidarOrdenVentaEmbarqueComplemento();
                if (estadoorden == "Iniciado")
                {
                    MessageBox.Show("La Orden de Venta Origen tiene embarque Iniciado, Considere avisar a Embarques de Cualquier cambio y considere que los tiempos de Carga seran afectados", "Carga Iniciada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (estadoorden == "Finalizado")
                {
                    MessageBox.Show("La Orden de Venta Origen tiene embarque Terminado y cerrado. No se pueden realizar Modificaciones a los productos hasta que la orden sea abierta", "Carga Terminada", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtcompleped.Focus();
                    txtcompleped.SelectAll();
                    return;
                }
            }

            string pedidoorigen = "N";

            string pedidoCargado = "N";

            pedidoCargado = ValidaCarga();

            if (pedidoCargado != "N")
            {
                if (pedidoCargado == "C")
                {
                    MessageBox.Show("Error: EL PEDIDO ACTUAL YA ESTA EN INICIO DE CARGA, FAVOR DE AVISAR A EMBARQUES PARA CANCELARLO, DESCARGAR EL PEDIDO Y REALICE LA MODIFICACION NUEVAMENTE", "PEDIDO EN PROCESO DE CARGA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //txtcompleped.Text = pedidocomplemento;
                    return;
                }
                else if (pedidoCargado == "T")
                {
                    MessageBox.Show("Error: ESTE PEDIDO YA SE HA CERRADO EN LA CARGA, NO SE PUEDE CAMBIAR EL PEDIDO ORIGEN/COMPLEMENTO", "PEDIDO YA CARGADO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //txtcompleped.Text = pedidocomplemento;
                    return;
                }
            }

            if (txtcompleped.Text.Length > 0 && txtcompleped.Text != "0")
            {
                pedidoorigen = ValidaPedidoOrigen();
            }
            if (pedidoorigen != "N")
            {
                MessageBox.Show("Error: ESTE PEDIDO " + txtcompleped.Text + System.Environment.NewLine + "NO ES EL PEDIDO ORIGEN/PRINCIPAL, EL PEDIDO ESTA ASOCIADO A LA ORDEN: " + pedidoorigen, "PEDIDO NO ES ORIGEN/PRINCIPAL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtcompleped.Text = "";
                return;
            }

            if (txtcompleped.Text.Trim() == txtfolio.Text.Trim())
            {
                MessageBox.Show("Error: ESTE PEDIDO " + txtcompleped.Text + System.Environment.NewLine + " Es Igual Al Pedido Actual, El pedido Principal/Origen no puede ser el mismo que el pedido actual", "PEDIDO IGUAL AL PEDIDO PRINCIPAL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtcompleped.Text = "";
                return;
            }
        }

        public string ValidarOrdenVentaEmbarqueComplemento()
        {
            thisConnection.Open();
            string estatus = "", hora_ini = "", hora_fin = "", sts = "";
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "select hora_ini, hora_fin, sts from tb_mstr_embarque where emb_folio = '" + txtcompleped.Text + "' and emb_tipo = 'EXP'";
            reader2 = cmnd2.ExecuteReader();
            while (reader2.Read())
            {
                hora_ini = reader2.GetValue(0).ToString().Trim();
                hora_fin = reader2.GetValue(1).ToString().Trim();
                sts = reader2.GetValue(2).ToString().Trim();
            }
            reader2.Dispose();
            thisConnection.Close();

            if (hora_ini.Trim() == "")
            {
                estatus = "SinIniciar";
            }
            else
            {
                estatus = "Iniciado";
                if (hora_ini.Trim() != "--:--")
                {
                    if ((hora_fin.Trim() != "--:--") && (sts.Trim() == "T"))
                    {
                        estatus = "Finalizado";
                    }
                }
            }

            return estatus;
        }

        DataTable recibos = new DataTable("recibos");
        DataSet ds = new DataSet();
        public decimal peso_promedio(string prod, string tipo)
        {
            decimal mkilos = 0, prom = 0, env_peso = 0;
            string query = "";
            recibos.Rows.Clear();
            recibos.Columns.Clear();
            #region si el producto es de campo va a tb_det_trazabilidad
            if (tipo.Trim() == "PTC")
            {
                query = "select a.*,  b.prod_clave, b.rptd_peso_bruto, b.rptd_tara, b.rptd_cantidad, c.rpt_fecha, d.env_peso from tb_det_trazabilidad a, tb_det_recepcion_pt b, " +
                        "tb_mstr_recepcion_pt c, tb_cat_envases d where a.prod_clave = '" + prod + "' and A.pti_estatus_sur = '' and b.rpt_recibo = a.recibo and " +
                        "b.prod_clave = a.prod_clave and c.rpt_recibo = a.recibo and d.env_clave = b.env_clave and c.rpt_estatus = '' and a.tipo = 'PTC' " +
                        "and (c.rpt_tipo != 'TR' OR (c.rpt_tipo = 'TR' AND c.RPT_INVENTARIO = 'S')) order by recibo desc";
                da = new SqlDataAdapter(query, thisConnection);
                da.Fill(ds, "recibos");
                recibos = ds.Tables["recibos"];

                foreach (DataRow row in recibos.Rows)
                {
                    if (Convert.ToDecimal(row["rptd_cantidad"]) == 0)
                        continue;
                    mkilos = mkilos + ((Convert.ToDecimal(row["rptd_peso_bruto"].ToString().Trim()) - Convert.ToDecimal(row["rptd_tara"].ToString().Trim())) / (Convert.ToDecimal(row["rptd_cantidad"].ToString())));

                    //if ((row["prod_clave"].ToString().Trim() == "02002ML00") || (row["prod_clave"].ToString().Trim() == "02002BROFR") || (row["prod_clave"].ToString().Trim() == "02BRCO2025"))                    
                    //    //mkilos = mkilos + (Fn_PesoHielo(Convert.ToDecimal(8.5), Convert.ToDateTime(row["rpt_fecha"].ToString()), Convert.ToDateTime(DateTime.Now)) * Convert.ToDecimal(row["rptd_cantidad"].ToString()));
                    //    mkilos = mkilos + (Fn_PesoHielo(Convert.ToDecimal(8.5), Convert.ToDateTime(row["rpt_fecha"].ToString()), Convert.ToDateTime(DateTime.Now)));

                    //if (row["prod_clave"].ToString().Trim() == "02002BRHEB")
                    //    mkilos = mkilos + (Fn_PesoHielo(Convert.ToDecimal(4), Convert.ToDateTime(row["rpt_fecha"].ToString()), Convert.ToDateTime(DateTime.Now)));
                }

                if (recibos.Rows.Count != 0)
                    prom = mkilos / recibos.Rows.Count;
                else
                {
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select prod_peso from tb_cat_producto where prod_clave='" + prod + "'";
                    prom = Convert.ToDecimal(cmnd2.ExecuteScalar());
                }
            }
            #endregion

            #region si el producto es de campo va a tb_det_eti_final
            if (tipo.Trim() == "PTP")
            {
                query = "SELECT A.*, B.PROD_CLAVE,B.PROD_PESO_VAR,B.FODP_UNIDADES, c.prod_nombre, c.prod_presentacion, d.HRP_PESO_NETO,d.HRP_NUM_UNIDADES,d.hrp_fecha, e.env_peso " +
                        "FROM tb_det_eti_final A, tb_det_final_odp b, tb_cat_producto c, tb_hist_recepcion d, tb_cat_envases e " +
                        "WHERE A.cve_prod = '" + prod + "' AND A.ESTATUS_SUR = '' and b.ordp_folio = a.folio and b.prod_clave= a.cve_prod and c.prod_clave = a.cve_prod " +
                        "and d.hrp_recibo = a.folio and d.prod_clave = a.cve_prod and e.env_clave = c.prod_presentacion ORDER BY A.folio DESC";
                da = new SqlDataAdapter(query, thisConnection);
                da.Fill(ds, "recibos");
                recibos = ds.Tables["recibos"];

                foreach (DataRow row in recibos.Rows)
                {
                    if (Convert.ToDecimal(row["hrp_num_unidades"]) == 0)
                        continue;
                    env_peso = Convert.ToDecimal(row["env_peso"].ToString());
                    //if (Convert.ToDecimal(row["prod_peso_var"].ToString()) > 0)
                    //    env_peso = Convert.ToDecimal(row["prod_peso_var"].ToString()) * Convert.ToDecimal(row["hrp_num_unidades"].ToString());
                    //else
                    //{
                    //    env_peso = Convert.ToDecimal(row["env_peso"].ToString()) * Convert.ToDecimal(row["hrp_num_unidades"].ToString());
                    //    if (env_peso == 0)
                    //        env_peso = 1;
                    //}
                    //mkilos = mkilos + ((Convert.ToDecimal(row["hrp_peso_neto"].ToString()) / Convert.ToDecimal(row["hrp_num_unidades"].ToString())) + (env_peso * Convert.ToDecimal(row["hrp_num_unidades"].ToString())));
                    decimal hrp_peso_neto = Convert.ToDecimal(row["hrp_peso_neto"].ToString());
                    decimal hrp_num_unidades = Convert.ToDecimal(row["hrp_num_unidades"].ToString());
                    mkilos = mkilos + ((hrp_peso_neto / hrp_num_unidades) + env_peso);

                    if ((row["prod_clave"].ToString().Trim() == "02002ML00") || (row["prod_clave"].ToString().Trim() == "02002BROFR") || (row["prod_clave"].ToString().Trim() == "02BRCO2025"))
                        //mkilos = mkilos + (Fn_PesoHielo(Convert.ToDecimal(8.5), Convert.ToDateTime(row["hrp_fecha"].ToString()), Convert.ToDateTime(DateTime.Now)) * Convert.ToDecimal(row["hrp_num_unidades"].ToString()));
                        mkilos = mkilos + (Fn_PesoHielo(Convert.ToDecimal(8.5), Convert.ToDateTime(row["hrp_fecha"].ToString()), Convert.ToDateTime(DateTime.Now)));

                    if (row["prod_clave"].ToString().Trim() == "02002BRHEB")
                        mkilos = mkilos + (Fn_PesoHielo(Convert.ToDecimal(4), Convert.ToDateTime(row["hrp_fecha"].ToString()), Convert.ToDateTime(DateTime.Now)));
                }

                if (recibos.Rows.Count != 0)
                    prom = mkilos / recibos.Rows.Count;
                else
                {
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select prod_peso from tb_cat_producto where prod_clave='" + prod + "'";
                    prom = Convert.ToDecimal(cmnd2.ExecuteScalar());
                }
            }
            #endregion
            return prom;
        }

        private decimal TraePesoEst(string Cve)
        {
            decimal mKilos = 0, mPeso = 0;
            string Cadena = "SELECT C.PROD_NOMBRE, A.RECIBO, A.TARIMA, A.PTI_FECHA, A.LOTE, A.FECHA_CAD, A.ETIQUETA, A.SURTIDO, A.PROD_CLAVE, A.UBICACION, AA.*, AB.ENV_PESO, AC.RPT_FECHA  FROM TB_DET_TRAZABILIDAD A, " +
                     "tb_mstr_recepcion_pt B, tb_cat_producto C, tb_det_recepcion_pt AA, tb_cat_envases AB, TB_MSTR_RECEPCION_PT AC  WHERE AA.rptd_estatus != 'C' AND  AA.RPT_RECIBO = A.RECIBO AND " +
                     "AA.PROD_CLAVE = A.prod_clave AND A.PROD_CLAVE = '" + Cve + "' AND AA.ENV_CLAVE = AB.ENV_CLAVE AND AA.RPT_RECIBO = AC.RPT_RECIBO AND A.PTI_ESTATUS_SUR =  ' ' AND A.TIPO = 'PTC'  AND A.recibo = B.rpt_recibo AND " +
                     "A.PROD_CLAVE = C.PROD_CLAVE AND B.rpt_estatus = ' ' AND (B.rpt_tipo != 'TR' OR (B.rpt_tipo = 'TR' AND B.RPT_INVENTARIO = 'S')) ORDER BY PROD_NOMBRE,RECIBO,PTI_CLAVE ";
            DataSet ds = new DataSet();
            DataTable Info = new DataTable();
            da = new SqlDataAdapter(Cadena, thisConnection);
            da.Fill(ds, "Info");
            Info = ds.Tables["Info"];
            Int32 mCjs = 0;
            foreach (DataRow row in Info.Rows)
            {
                mKilos = mKilos + ((Convert.ToDecimal(row["rptd_peso_bruto"]) - Convert.ToDecimal(row["rptd_tara"])) / Convert.ToDecimal(row["rptd_cantidad"])) * Convert.ToDecimal(Convert.ToInt32(row["ETIQUETA"]) - Convert.ToInt32(row["SURTIDO"]));
                DateTime mF = DateTime.Now;  //Convert.ToDateTime(Program.MyGlobal.PubFecEmb);
                if (Convert.ToString(row["PROD_CLAVE"]) == "02002ML00" || Convert.ToString(row["PROD_CLAVE"]) == "02002BROFR" || Convert.ToString(row["PROD_CLAVE"]) == "02BRCO2025")
                    mKilos = mKilos + (Fn_PesoHielo(Convert.ToDecimal(8.5), Convert.ToDateTime(row["RPT_FECHA"]), mF) * Convert.ToInt32(row["ETIQUETA"]) - Convert.ToInt32(row["SURTIDO"]));
                if (row["PROD_CLAVE"].ToString() == "02002BRHEB")
                    mKilos = mKilos + Fn_PesoHielo(Convert.ToDecimal(4), Convert.ToDateTime(row["RPT_FECHA"]), mF);
                mCjs += Convert.ToInt32(Convert.ToInt32(row["ETIQUETA"]) - Convert.ToInt32(row["SURTIDO"]));
            }
            if (mKilos > 0 && mCjs > 0)
                mPeso = mKilos / mCjs;
            Info.Rows.Clear();
            Cadena = "SELECT A.PROD_NOMBRE, B.FOLIO, B.TARIMA, B.FECHA, B.NUM_LOTE, B.NUM_CAJAS, B.CAJAS_SUR, B.CVE_PROD, B.UBICACION, b.fechacad, AB.PROD_CLAVE, AC.PROD_NOMBRE, AB.PROD_PESO_VAR,AB.FODP_UNIDADES," +
                     "AC.PROD_PRESENTACION,AD.ENV_PESO,AE.HRP_PESO_NETO,AE.HRP_NUM_UNIDADES,AE.hrp_fecha FROM TB_DET_ETI_FINAL B, TB_CAT_PRODUCTO A, TB_DET_FINAL_ODP AB, TB_CAT_PRODUCTO AC,tb_cat_envases AD, " +
                     "TB_HIST_RECEPCION AE  WHERE   AB.PROD_CLAVE = A.PROD_CLAVE AND AB.ORDP_FOLIO = B.FOLIO AND AB.PROD_CLAVE = AC.PROD_CLAVE AND AC.PROD_PRESENTACION = AD.ENV_CLAVE AND AB.ORDP_FOLIO = AE.hrp_recibo " +
                     "AND B.CVE_PROD = '" + Cve + "' AND AB.PROD_CLAVE = AE.PROD_CLAVE AND B.ESTATUS_SUR =  ' ' AND B.CVE_PROD = A.PROD_CLAVE AND B.ETIQUETA = 'S' ORDER BY A.PROD_NOMBRE,B.FOLIO,B.TARIMA ";
            ds = new DataSet();
            da = new SqlDataAdapter(Cadena, thisConnection);
            da.Fill(ds, "Info");
            Info = ds.Tables["Info"];
            mKilos = 0; mCjs = 0;
            foreach (DataRow row in Info.Rows)
            {
                int MCant = Convert.ToInt32(row["NUM_CAJAS"]) - Convert.ToInt32(row["CAJAS_SUR"]);
                decimal Mpe = 0;
                if (Convert.ToDecimal(row["PROD_PESO_VAR"]) > 0)
                    Mpe = Convert.ToDecimal(row["PROD_PESO_VAR"]) * MCant;
                else
                    Mpe = Convert.ToDecimal(row["ENV_PESO"]) * MCant;

                //foreach (DataRow rowx in PesoProd.Select("Prod_Clave ='" + row["PROD_CLAVE"].ToString() + "'"))
                //    Mpe = (Convert.ToDecimal(rowx["env_peso"]) == 0) ? 1 : Convert.ToDecimal(rowx["env_peso"]);

                decimal Muni = Convert.ToDecimal(row["HRP_NUM_UNIDADES"]);
                mKilos = mKilos + ((Convert.ToDecimal(row["HRP_PESO_NETO"]) / Muni) * MCant) + (Mpe * MCant); // +Mpe;
                //MKILOS = MKILOS + ((HRP_PESO_NETO / HRP_NUM_UNIDADES) * D) + (var_dec_env_peso * D)
                if (row["PROD_CLAVE"].ToString().Trim() == "02002ML00" || row["PROD_CLAVE"].ToString() == "02002BROFR" || row["PROD_CLAVE"].ToString() == "02BRCO2025")
                    mKilos = mKilos + (Fn_PesoHielo(Convert.ToDecimal(8.5), Convert.ToDateTime(row["HRP_FECHA"]), DateTime.Now) * MCant);
                if (row["PROD_CLAVE"].ToString() == "02002BRHEB")
                    mKilos = mKilos + Fn_PesoHielo(Convert.ToDecimal(4), Convert.ToDateTime(row["HRP_FECHA"]), DateTime.Now);
                mCjs += MCant;
            }
            if (mKilos > 0 && mCjs > 0)
                mPeso += (mKilos / mCjs);
            return mPeso;
        }

        public decimal Fn_PesoHielo(decimal peso, DateTime FecRec, DateTime FecEmb)
        {
            int MDIA = FecEmb.Subtract(FecRec).Days;
            decimal Mpeso = 0;
            switch (MDIA)
            {
                case 0:
                    Mpeso = peso;
                    break;
                case 1:
                    Mpeso = (peso * 85) / 100;
                    break;
                case 2:
                    Mpeso = (peso * 75) / 100;
                    break;
                case 3:
                    Mpeso = (peso * 50) / 100;
                    break;
                case 4:
                    Mpeso = (peso * 35) / 100;
                    break;
                case 5:
                    Mpeso = (peso * 20) / 100;
                    break;
                case 6:
                    Mpeso = (peso * 10) / 100;
                    break;
                default:
                    Mpeso = peso;
                    break;
            }
            return Mpeso;
        }

        private void DGV2_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (DGV2.Columns[e.ColumnIndex].Name == "cajas" && opcion == 3) // modificacion
                CantMod = Convert.ToInt32(DGV2.CurrentRow.Cells["cajas"].Value);
        }

        private void ValidaSplit(string Mped, string ProdCve, int mCant, string NomProd)
        {
            int total = 0;
            thisConnection.Open();
            string Cadena = "SELECT isnull(SUM(cajas), 0) as cajas, nom_capsplit FROM tb_det_split WHERE emb_folio = '0" + Mped + "' AND prod_clave = '" + ProdCve + "' AND estatus = 'A'  GROUP BY nom_capsplit";
            SqlDataAdapter da = new SqlDataAdapter(Cadena, thisConnection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Splits");
            DataTable Split = ds.Tables["Splits"];
            int mTcjs = 0;
            string mRes = "";
            foreach (DataRow row in Split.Rows)
            {
                mTcjs += Convert.ToInt32(row["cajas"]);
                mRes += row["nom_capsplit"];
            }
            Cadena = "SELECT isnull(SUM(cajas), 0) as cajas FROM  tb_det_embarque WHERE emb_folio = '0" + Mped + "' AND prod_clave = '" + ProdCve + "' AND estatus != 'C'";
            da = new SqlDataAdapter(Cadena, thisConnection);
            ds = new DataSet();
            da.Fill(ds, "cargado");
            DataTable cargado = ds.Tables["cargado"];
            foreach (DataRow row in cargado.Rows)
            {
                total += Convert.ToInt32(row["cajas"]);
            }
            if ((mTcjs + total) > 0)
            { // ENCONTRO SPLITS ARMANDOSE
                if (mCant < (mTcjs + total))
                {
                    if (mTcjs > 0)
                    {
                        MessageBox.Show("Ya se tienen Splits Armados para Este producto: " + NomProd + System.Environment.NewLine +
                                "Si Modifica la cantidad necesita avisar a embarques ya que tienen que retrabajar en modificar el split" + System.Environment.NewLine +
                                "Esta Armando el Split: " + mRes, "PRODUCTO ARMADO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    if (total > 0)
                    {
                        MessageBox.Show("Ya se Cargo Este producto: " + NomProd + System.Environment.NewLine +
                                      "Se han Cargado " + total + " Cajas, Para modificar la cantidad necesita avisar a embarques ya que tienen que bajar el producto" + System.Environment.NewLine +
                                      "CONSIDERE QUE EL EMBARQUE VA A RETRASARSE En su Salida", "PRODUCTO YA CARGADO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    DGV2.CurrentRow.Cells["cajas"].Value = CantMod;
                }// si la cantidad modificada es menor que lo que ya se tiene en el Split notificarle al ejecutivo de ventas
            }
            thisConnection.Close();
        }

        private string ValidaSplitBorrado(string Mped, string ProdCve, int mCant, string NomProd)
        {
            string encontrado = "0";
            thisConnection.Open();
            string Cadena = "SELECT * FROM tb_det_split WHERE emb_folio = '0" + Mped + "' AND prod_clave = '" + ProdCve + "' AND estatus != 'C'";
            SqlDataAdapter da = new SqlDataAdapter(Cadena, thisConnection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Splits");
            DataTable Split = ds.Tables["Splits"];
            int mTcjs = 0;
            int total = 0;
            string mRes = "";
            foreach (DataRow row in Split.Rows)
            {
                mTcjs += Convert.ToInt32(row["cajas"]);
                mRes += row["nom_capsplit"];
            }
            Cadena = "SELECT isnull(SUM(cajas), 0) as cajas FROM  tb_det_embarque WHERE emb_folio = '0" + Mped + "' AND prod_clave = '" + ProdCve + "' AND estatus != 'C'";
            da = new SqlDataAdapter(Cadena, thisConnection);
            ds = new DataSet();
            da.Fill(ds, "cargado");
            DataTable cargado = ds.Tables["cargado"];
            foreach (DataRow row in cargado.Rows)
            {
                total += Convert.ToInt32(row["cajas"]);
            }


            if ((mTcjs + total) > 0)
            { // ENCONTRO SPLITS ARMANDOSE
                if (mCant < (mTcjs + total))
                {
                    encontrado = "1";
                    if (mTcjs > 0)
                    {
                        MessageBox.Show("Ya se tienen Splits Armados para Este producto: " + NomProd + System.Environment.NewLine +
                                "Si Modifica la cantidad necesita avisar a embarques ya que tienen que retrabajar en modificar el split" + System.Environment.NewLine +
                                "Esta Armando el Split: " + mRes, "PRODUCTO ARMADO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    if (total > 0)
                    {
                        MessageBox.Show("Ya se Cargo Este producto: " + NomProd + System.Environment.NewLine +
                                      "Se han Cargado " + total + " Cajas, Para modificar la cantidad necesita avisar a embarques ya que tienen que bajar el producto" + System.Environment.NewLine +
                                      "CONSIDERE QUE EL EMBARQUE VA A RETRASARSE En su Salida", "PRODUCTO YA CARGADO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    //DGV2.CurrentRow.Cells["cajas"].Value = CantMod;
                }// si la cantidad modificada es menor que lo que ya se tiene en el Split notificarle al ejecutivo de ventas
            }
            thisConnection.Close();
            return encontrado;
        }

        private void DGV2_CellLeave(object sender, DataGridViewCellEventArgs e)
        {


        }

        public string ValidarOrdenVentaEmbarque()
        {
            thisConnection.Open();
            string estatus = "", hora_ini = "", hora_fin = "", sts = "";
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "select hora_ini, hora_fin, sts from tb_mstr_embarque where emb_folio = '0" + txtfolio.Text + "' and emb_tipo = 'EXP'";
            reader2 = cmnd2.ExecuteReader();
            while (reader2.Read())
            {
                hora_ini = reader2.GetValue(0).ToString().Trim();
                hora_fin = reader2.GetValue(1).ToString().Trim();
                sts = reader2.GetValue(2).ToString().Trim();
            }
            reader2.Dispose();
            thisConnection.Close();

            if (hora_ini.Trim() == "")
            {
                estatus = "SinIniciar";
            }
            else
            {
                estatus = "Iniciado";
                if (hora_ini.Trim() != "--:--")
                {
                    if ((hora_fin.Trim() != "--:--") && (sts.Trim() == "T"))
                    {
                        estatus = "Finalizado";
                    }
                }
            }

            return estatus;
        }


        #region METODOS PARA VALIDAR CAMBIOS EN EL PEDIDO
        public string ValidarCambiosEmbarqueOrdenV2()
        {
            DataTable detallePedidoBD = new DataTable();

            // 1. Uso de 'using' para asegurar que la conexión y el comando se cierren correctamente (Evita Fugas de memoria y bloqueos en la base de datos)
            using (SqlCommand cmd = new SqlCommand("SELECT prod_clave, pdn_num_unidades FROM tb_det_pedidos WHERE pdn_folio = @folio AND pdn_tipo = 'EXP'", thisConnection))
            {
                // 2. Uso de Parámetros de seguridad (previene SQL Injection; Manipulación de la Información) 
                cmd.Parameters.AddWithValue("@folio", txtfolio.Text.Trim());

                try
                {
                    thisConnection.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        detallePedidoBD.Load(reader);
                    }
                }
                finally
                {
                    thisConnection.Close();
                }
            }

            // 3. Verificación rápida: Si el conteo de filas es distinto, ya sabemos que hay cambios
            if (DGV2.Rows.Count != detallePedidoBD.Rows.Count)
                return "SI";

            // 4. Diccionario para búsqueda ultra rápida (Optimización de rendimiento reemplazando los ciclos for e if anidados; Evitando complejidad O(n²) Por ejemplo si tenemos n productos el codigo original hacia hasta n² comparaciones) 
            // Guardamos la clave del producto y su cantidad
            var dictBD = detallePedidoBD.AsEnumerable()
                .ToDictionary(
                    row => row["prod_clave"].ToString().Trim(),
                    row => Convert.ToInt32(row["pdn_num_unidades"].ToString().Replace(".000", ""))
                );

            // 5. Comparación lógica
            foreach (DataGridViewRow filaGrid in DGV2.Rows)
            {
                string claveProd = filaGrid.Cells["clave_producto"].Value?.ToString().Trim();
                decimal valorDecimal = Convert.ToDecimal(filaGrid.Cells["cajas"].Value ?? 0);
                int cantidadGrid = (int)Math.Round(valorDecimal);
                //int cantidadGrid = Convert.ToInt32(filaGrid.Cells["cajas"].Value);

                // ¿El producto del Grid existe en la BD?
                if (dictBD.TryGetValue(claveProd, out int cantidadBD))
                {
                    // Si existe, ¿la cantidad es diferente?
                    if (cantidadGrid != cantidadBD) return "SI";
                }
                else
                {
                    // Si no existe el producto en la BD, es un cambio
                    return "SI";
                }
            }

            return "NO";
        }
        public string ValidarCambiosEmbarqueOrden()
        {
            thisConnection.Open();
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "select prod_clave, pdn_num_unidades FROM tb_det_pedidos where pdn_folio = '" + txtfolio.Text + "' and pdn_tipo = 'EXP'";
            reader2 = cmnd2.ExecuteReader();
            DataTable detallepedidoactual = new DataTable();
            detallepedidoactual.Load(reader2);
            reader2.Dispose();
            thisConnection.Close();

            string estadocambio = "NO";

            for (int i = 0; i < DGV2.Rows.Count; i++)
            {
                string encontrado = "0";
                foreach (DataRow row in detallepedidoactual.Select("prod_clave = '" + DGV2.Rows[i].Cells["clave_producto"].Value.ToString().Trim() + "'"))
                {
                    encontrado = "1";
                    int xc = Convert.ToInt32(DGV2.Rows[i].Cells["cajas"].Value);
                    int xs = Convert.ToInt32(row["pdn_num_unidades"].ToString().Replace(".000", ""));
                    if (Convert.ToInt32(DGV2.Rows[i].Cells["cajas"].Value) != Convert.ToInt32(row["pdn_num_unidades"].ToString().Replace(".000", "")))
                    {
                        estadocambio = "SI";
                    }
                }

                if (encontrado == "0")
                {
                    estadocambio = "SI";
                }
            }

            foreach (DataRow row in detallepedidoactual.Rows)
            {
                string encontradootro = "0";
                for (int i = 0; i < DGV2.Rows.Count; i++)
                {
                    if (row["prod_clave"].ToString().Trim() == DGV2.Rows[i].Cells["clave_producto"].Value.ToString().Trim())
                    {
                        encontradootro = "1";
                    }
                }
                string xa = row["prod_clave"].ToString();

                if (encontradootro == "0")
                {
                    estadocambio = "SI";
                }
            }
            return estadocambio;
        }
        #endregion

        private void CreaTabla()
        {
            Fraccion.Columns.Add("Cve", typeof(String));
            Fraccion.Columns.Add("Nombre", typeof(String));
            Fraccion.Columns.Add("Fraccion", typeof(String));
            Fraccion.Columns.Add("Aduana", typeof(String));
            Fraccion.Columns.Add("Cajas", typeof(decimal));
            Fraccion.Columns.Add("Peso", typeof(decimal));
            Fraccion.Columns.Add("Cant", typeof(decimal));
            Fraccion.Columns.Add("PreUni", typeof(decimal));
            Fraccion.Columns.Add("Importe", typeof(decimal));
            Fraccion.Columns.Add("Total", typeof(decimal));
        }

        private void BtnComple_Click(object sender, EventArgs e)
        {
            FrmCompleExp FrmCom = new FrmCompleExp();
            FrmCom.Info = Fraccion;
            FrmCom.Folio = txtfolio.Text;
            if (CveSubCli.Trim().Length > 0 && txtcliente.Text.Trim() == "GABOP")
            {
                FrmCom.Destino = cbdestino.Text.Trim();
                foreach (DataRow row in subcnte.Select("cnte_clave = 'GABOP' AND ENTR_CLAVE= '" + CveSubCli + "'"))
                    FrmCom.Domicilio = row["ENTR_DIRECCION"].ToString().Trim() + ", " + row["ENTR_CIUDAD"].ToString().Trim() + " " +
                                       row["ENTR_ESTADO"].ToString().Trim() + " " + row["ENTR_CP"].ToString().Trim() + " TAX ID: " + row["ENTR_RFC"].ToString().Trim();
            }
            else
            {
                FrmCom.Destino = CBCliente.Text.Trim();
                FrmCom.Domicilio = DomiCli(CBCliente.Text.Trim());
                //FrmCom.Domicilio = direccion.Text.Trim() + ", " + ((colonia.Text.Trim() == "N/A") ? " " : colonia.Text.Trim() + " ") +
                //                   ciudad.Text.Trim() + " " + estado.Text.Trim() + " " + LblCodPos.Text.Trim();
                //entr_direccion, entr_colonia, entr_ciudad, entr_estado, entr_cp
            }
            FrmCom.ShowDialog(this);
        }

        private string DomiCli(string Cve)
        {
            string Domi = "";
            thisConnection.Open();
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "select cnte_nombre, cnte_direccion, cnte_colonia, cnte_ciudad, cnte_estado, cnte_telefono, cnte_entregas, cnte_rfc, cnte_entregas, cnte_condicion, cnte_cp from tb_cat_cliente where cnte_clave = '" + txtcliente.Text + "'";
            reader2 = cmnd2.ExecuteReader();
            while (reader2.Read())
            {
                string direccion = reader2.GetSqlString(1).ToString().Trim();
                string colonia = reader2.GetSqlString(2).ToString().Trim();
                string ciudad = reader2.GetSqlString(3).ToString().Trim();
                string estado = reader2.GetSqlString(4).ToString().Trim();
                string Codpos = reader2.GetSqlString(10).ToString().Trim();
                Domi = direccion + ", " + ((colonia == "N/A") ? " " : colonia + " ") +
                        ciudad + " " + estado + " " + Codpos;
            }
            reader2.Close();
            thisConnection.Close();
            return Domi;
        }

        private void btncopiarpedido_Click(object sender, EventArgs e)
        {
            lbpedido.Visible = true;
            txtpedido.Visible = true;
            txtpedido.Enabled = true;
            TipoPed = "C"; // Copiar
            txtpedido.Focus();
        }

        private void txtpedido_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                if (TipoPed == "C")
                    CopiarPedido();
                else
                    PedidoTomate();
            }
        }

        private void CopiarPedido()
        {
            if (MessageBox.Show("¿Esta seguro de importar los productos del pedido solicitado?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //limpia grid de tarimas
                DataTable DT = (DataTable)DGV2.DataSource;
                if (DT != null)
                    DT.Clear();

                try
                {
                    thisConnection.Open();
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "SELECT tdp.lin_clave, tdp.prod_clave, tcpr.prod_nombre, tdp.pdn_num_unidades as cantidad, tdp.pdn_pedimento, " + //cast(tdp.pdn_num_unidades as int) as cantidad
                                        "tdp.pdn_fechap, tdp.pdn_aduana, " +
                                        "tdp.pdn_precio_usd, tdp.pdn_cantxcaja, tdp.pdn_pre_cad, cast((tdp.pdn_num_unidades * tdp.pdn_precio_usd) as decimal(9,2)) AS total, " +
                                        "cast((tdp.pdn_num_unidades * tdp.pdn_precio_usd) as decimal(9,2)) AS total, pdn_observaciones as observaciones, " +
                                        "tcpr.prod_tipo as tipo, tdp.pdn_tipotar FROM tb_det_pedidos tdp, tb_cat_producto tcpr " +
                                        "WHERE tdp.pdn_folio = '" + txtpedido.Text + "' and tdp.pdn_tipo = 'EXP' AND tcpr.prod_clave = tdp.prod_clave ORDER BY tdp.nrenglon ASC";
                    reader2 = cmnd2.ExecuteReader();
                    if (reader2.HasRows == false)
                    {
                        MessageBox.Show("Aviso No. Pedido no Existe!!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    while (reader2.Read())
                    {
                        DataRow drc = tmgrdped.NewRow();
                        drc["linea"] = reader2.GetSqlString(0).ToString();
                        drc["producto"] = reader2.GetSqlString(1).ToString();
                        drc["nombre"] = reader2.GetSqlString(2).ToString();
                        drc["cantidad"] = Convert.ToInt32(reader2["cantidad"]).ToString("####");
                        drc["pedimento"] = reader2.GetSqlString(4).ToString();
                        //if (reader2.GetSqlString(5).ToString() == "Null")
                        //    drc["fecha"] = " ";

                        //else
                        //    drc["fecha"] = reader2.GetSqlString(5).ToString();
                        //drc["aduana"] = reader2.GetSqlString(6).ToString();
                        drc["precio"] = reader2.GetSqlDecimal(7).ToString();
                        drc["clave_tarima"] = reader2["pdn_tipotar"].ToString();
                        foreach (DataRow dr in tarimas.Select("emp_clave = '" + reader2["pdn_tipotar"].ToString() + "'"))
                        {
                            drc["nom_tarima"] = Convert.ToString(dr["emp_nombre"].ToString().Trim());
                        }
                        //cmnd3 = thisConnection.CreateCommand();
                        //cmnd3.CommandText = "select emp_nombre from tb_cat_empaques where emp_clave = '" + reader2["pdn_tipotar"].ToString() + "'";
                        //drc["nom_tarima"] = Convert.ToString(cmnd3.ExecuteScalar()).Trim();
                        //if (reader2.GetSqlString(9).ToString().Trim() != "")
                        //    drc["precio_uni"] = reader2.GetSqlString(9).ToString();
                        //else
                        //    drc["precio_uni"] = 0;
                        drc["total"] = reader2.GetSqlDecimal(10).ToString();
                        drc["observaciones"] = reader2.GetSqlString(12).ToString();
                        //drc["tipo"] = reader2.GetValue(13).ToString().Trim();
                        tmgrdped.Rows.Add(drc);
                    }
                    reader2.Dispose();
                    thisConnection.Close();
                    DGV2.DataSource = tmgrdped;
                    //foreach (DataRow dr in tmgrdped.Rows)
                    //{
                    //    decimal canti = Convert.ToDecimal(dr["cantidad"].ToString());
                    //    decimal pre = Convert.ToDecimal(dr["precio"].ToString());
                    //    decimal pre_uni = Convert.ToDecimal(dr["precio_uni"].ToString());
                    //    decimal tot = Convert.ToDecimal(dr["total"].ToString());
                    //    DGV2.Rows.Add(Convert.ToString(dr["linea"].ToString().Trim()), Convert.ToString(dr["producto"].ToString().Trim()), Convert.ToString(dr["nombre"].ToString().Trim()), canti.ToString("###,##0"), Convert.ToString(dr["pedimento"].ToString().Trim()), Convert.ToString(dr["fecha"].ToString().Trim()), Convert.ToString(dr["aduana"].ToString().Trim()), pre.ToString("###,##0.00"), tot.ToString("###,###,##0.00"), Convert.ToString(dr["unidad"].ToString().Trim()), pre.ToString("###,###,##0.00"), Convert.ToString(dr["observaciones"].ToString().Trim()));
                    //}
                    //this.DGV2.Columns["cajas"].DefaultCellStyle.Format = "0";
                    //this.DGV2.Columns[6].DefaultCellStyle.Format = "0.00";
                    //this.DGV2.Columns[11].DefaultCellStyle.Format = "0.00";
                    //this.DGV2.Columns[12].DefaultCellStyle.Format = "0.00";
                    // this.DGV2.Columns[14].DefaultCellStyle.Format = "0.00";
                    DGV2.Update();
                    DGV2.Enabled = true;
                    for (int i = 0; i < DGV2.Rows.Count; i++)
                    {
                        this.DGV2.Rows[i].HeaderCell.Value = (i + 1).ToString();
                    }
                    DGV2.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader);
                }
                catch (SqlException ex)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "sisgab", ex.ToString());
                    MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                catch (Exception ex1)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "sisgab", ex1.ToString());
                    MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }

        private void BtnPedTom_Click(object sender, EventArgs e)
        {
            lbpedido.Visible = true;
            txtpedido.Visible = true;
            txtpedido.Enabled = true;
            TipoPed = "T"; // Pedido del Tomate
            txtpedido.Focus();
        }

        private void PedidoTomate()
        {
            if (MessageBox.Show("¿Esta seguro de importar los productos del pedido solicitado?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //limpia grid de tarimas
                DataTable DT = (DataTable)DGV2.DataSource;
                if (DT != null)
                    DT.Clear();

                try
                {
                    thisConnection.Open();
                    thisConnecTomate.Open();
                    cmnd2 = thisConnecTomate.CreateCommand();
                    cmnd2.CommandText = "SELECT tdp.lin_clave, tdp.prod_clave, tcpr.prod_nombre, tdp.pdn_num_unidades as cantidad, tdp.pdn_pedimento, " + //cast(tdp.pdn_num_unidades as int) as cantidad
                                        "tdp.pdn_fechap, tdp.pdn_aduana, " +
                                        "tdp.pdn_precio_usd, tdp.pdn_cantxcaja, tdp.pdn_pre_cad, cast((tdp.pdn_num_unidades * tdp.pdn_precio_usd) as decimal(9,2)) AS total, " +
                                        "pdn_observaciones as observaciones, " +
                                        "tcpr.prod_tipo as tipo, tcpr.prod_tarimas FROM tb_det_pedidos tdp, tb_cat_producto tcpr " +
                                        "WHERE tdp.pdn_folio = '" + txtpedido.Text + "' and tdp.pdn_tipo = 'NAL' AND tcpr.prod_clave = tdp.prod_clave ORDER BY tdp.nrenglon ASC";
                    cmnd2.CommandTimeout = 0;
                    reader2 = cmnd2.ExecuteReader();
                    if (reader2.HasRows == false)
                    {
                        MessageBox.Show("Aviso No. Pedido no Existe!!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    while (reader2.Read())
                    {
                        cmnd3 = thisConnection.CreateCommand();
                        if (CBmoneda.SelectedIndex == 0)
                            cmnd3.CommandText = "SELECT tdp.pdn_precio_usd FROM tb_mstr_pedidos_exp tmpn, tb_det_pedidos tdp WHERE tmpn.cnte_clave='" + txtcliente.Text + "'  AND " +
                                                "tdp.prod_clave='" + reader2.GetSqlString(1).ToString() + "' and tdp.pdn_tipo = 'EXP' AND tmpn.pdn_folio = tdp.pdn_folio ORDER BY tmpn.pdn_fecha DESC";

                        if (CBmoneda.SelectedIndex == 1 || CBmoneda.SelectedIndex == 2)
                            cmnd3.CommandText = "SELECT tdp.pdn_precio_usd FROM tb_mstr_pedidos_exp tmpn, tb_det_pedidos tdp WHERE tmpn.cnte_clave='" + txtcliente.Text + "'  AND " +
                                                "tdp.prod_clave='" + reader2.GetSqlString(1).ToString() + "' and tdp.pdn_tipo = 'EXP' AND tmpn.pdn_folio = tdp.pdn_folio ORDER BY tmpn.pdn_fecha DESC";

                        reader3 = cmnd3.ExecuteReader();
                        precio = 0;
                        while (reader3.Read())
                        {
                            precio = reader3.GetSqlDecimal(0).Value;
                            encontro = true;
                            break;
                        }
                        reader3.Close();
                        DataRow drc = tmgrdped.NewRow();
                        drc["linea"] = reader2.GetSqlString(0).ToString();
                        drc["producto"] = reader2.GetSqlString(1).ToString();
                        drc["nombre"] = reader2.GetSqlString(2).ToString();
                        drc["cantidad"] = Convert.ToInt32(reader2["cantidad"]).ToString("####");
                        drc["pedimento"] = reader2.GetSqlString(4).ToString();
                        //if (reader2.GetSqlString(5).ToString() == "Null")
                        //    drc["fecha"] = " ";

                        //else
                        //    drc["fecha"] = reader2.GetSqlString(5).ToString();
                        //drc["aduana"] = reader2.GetSqlString(6).ToString();
                        drc["precio"] = precio; //reader2.GetSqlDecimal(7).ToString();
                        drc["clave_tarima"] = reader2["prod_tarimas"].ToString();
                        foreach (DataRow dr in tarimas.Select("emp_clave = '" + reader2["prod_tarimas"].ToString() + "'"))
                        {
                            drc["nom_tarima"] = Convert.ToString(dr["emp_nombre"].ToString().Trim());
                        }
                        //cmnd3 = thisConnection.CreateCommand();
                        //cmnd3.CommandText = "select emp_nombre from tb_cat_empaques where emp_clave = '" + reader2["pdn_tipotar"].ToString() + "'";
                        //drc["nom_tarima"] = Convert.ToString(cmnd3.ExecuteScalar()).Trim();
                        //if (reader2.GetSqlString(9).ToString().Trim() != "")
                        //    drc["precio_uni"] = reader2.GetSqlString(9).ToString();
                        //else
                        //    drc["precio_uni"] = 0;
                        drc["total"] = Convert.ToInt32(reader2["cantidad"]) * precio;// reader2.GetSqlDecimal(10).ToString();
                        drc["observaciones"] = ""; //reader2.GetSqlString(12).ToString();
                        //drc["tipo"] = reader2.GetValue(13).ToString().Trim();
                        tmgrdped.Rows.Add(drc);
                    }
                    reader2.Dispose();
                    thisConnection.Close();
                    thisConnecTomate.Close();
                    DGV2.DataSource = tmgrdped;
                    //foreach (DataRow dr in tmgrdped.Rows)
                    //{
                    //    decimal canti = Convert.ToDecimal(dr["cantidad"].ToString());
                    //    decimal pre = Convert.ToDecimal(dr["precio"].ToString());
                    //    decimal pre_uni = Convert.ToDecimal(dr["precio_uni"].ToString());
                    //    decimal tot = Convert.ToDecimal(dr["total"].ToString());
                    //    DGV2.Rows.Add(Convert.ToString(dr["linea"].ToString().Trim()), Convert.ToString(dr["producto"].ToString().Trim()), Convert.ToString(dr["nombre"].ToString().Trim()), canti.ToString("###,##0"), Convert.ToString(dr["pedimento"].ToString().Trim()), Convert.ToString(dr["fecha"].ToString().Trim()), Convert.ToString(dr["aduana"].ToString().Trim()), pre.ToString("###,##0.00"), tot.ToString("###,###,##0.00"), Convert.ToString(dr["unidad"].ToString().Trim()), pre.ToString("###,###,##0.00"), Convert.ToString(dr["observaciones"].ToString().Trim()));
                    //}
                    //this.DGV2.Columns["cajas"].DefaultCellStyle.Format = "0";
                    //this.DGV2.Columns[6].DefaultCellStyle.Format = "0.00";
                    //this.DGV2.Columns[11].DefaultCellStyle.Format = "0.00";
                    //this.DGV2.Columns[12].DefaultCellStyle.Format = "0.00";
                    // this.DGV2.Columns[14].DefaultCellStyle.Format = "0.00";
                    DGV2.Update();
                    DGV2.Enabled = true;
                    for (int i = 0; i < DGV2.Rows.Count; i++)
                    {
                        this.DGV2.Rows[i].HeaderCell.Value = (i + 1).ToString();
                    }
                    DGV2.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader);
                }
                catch (SqlException ex)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "sisgab", ex.ToString());
                    MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                catch (Exception ex1)
                {
                    thisConnection.Close();
                    Utilerias.Class1.SendMail("sistemas@mrlucky.com.mx", "sistemas", "sisgab", ex1.ToString());
                    MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }
    }
}

