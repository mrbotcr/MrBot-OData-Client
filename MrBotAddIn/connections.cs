using Microsoft.Data.Edm;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Simple.OData.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MrBotAddIn
{
    public partial class connections : Form
    {
        funcionesEspeciales fe = new funcionesEspeciales();
        ODataClient client;
        conexionesOData conexionSeleccionada;
        Ribbon1 evente = new Ribbon1();
        public connections()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*
                DELETE CONNECTION

                We obtain the list of connections of the properties of the project and convert them into a list.
                The connection with the name to be deleted is searched and removed from the list.
                The list is updated in the properties and changes are saved.
            */
            List<conexionesOData> lista = JsonConvert.DeserializeObject<List<conexionesOData>>(MrBotAddIn.Properties.Settings.Default.jsonDeConexiones);
            if (lista != null)
            {
                foreach (conexionesOData connect in lista)
                {
                    if (comboBox1.Text == connect.Name)
                    {
                        lista.Remove(connect);
                        break;
                    }
                }
            }
            MrBotAddIn.Properties.Settings.Default.jsonDeConexiones = JsonConvert.SerializeObject(lista);
            MrBotAddIn.Properties.Settings.Default.Save();
            cargarConexiones();
        }

        private void connections_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            cargarConexiones();            
        }

        public void cargarConexiones()
        {
            /*
                We obtain the list of connections of the properties of the project, 
                convert them into a list and load them in the comboboxes.
            */
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            List<conexionesOData> lista = JsonConvert.DeserializeObject<List<conexionesOData>>(MrBotAddIn.Properties.Settings.Default.jsonDeConexiones);
            if (lista != null)
            {
                foreach (conexionesOData connect in lista)
                {
                    comboBox1.Items.Add(connect.Name);
                }
            }
            comboBox1.Items.Add("New");
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*
                Every time the connection combobox changes an action is made.
                If the selected option is "New" then the interface to create a new connection is opened.
                If the option is one of the connections already created, it proceeds to obtain the tables it has. And they are loaded in the Tables combobox.
            */

            if (comboBox1.Text == "New")
            {
                nuevaConexion nuevaConexion = new nuevaConexion();
                nuevaConexion.Show();

                this.Close();
            }else
            {
                comboBox2.Text = "";
                List<conexionesOData> lista = JsonConvert.DeserializeObject<List<conexionesOData>>(MrBotAddIn.Properties.Settings.Default.jsonDeConexiones);
                conexionSeleccionada = lista.Find(x => x.Name == comboBox1.Text);
                client = new ODataClient(conexionSeleccionada.Url);
                llamada(client);
                

            }
        }

        public async void llamada(ODataClient client)
        {
            comboBox2.Items.Clear();
            IEdmModel metadata = await client.GetMetadataAsync<IEdmModel>();
            
            var entityTypes = metadata.SchemaElements.OfType<IEdmEntityType>().ToArray();
            foreach (var type in entityTypes)
            {
                comboBox2.Items.Add(type.Name);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //It is checked that a valid table is selected in the table combobox
            if (comboBox2.Text != "")
            {
                //A worksheet is created to show the data and assign the necessary variables
                Worksheet activeSheet = (Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                //Set the name of the active sheet
                activeSheet.Name = comboBox2.Text;
                /*
                    It is validated if the variables were already created in the properties of the sheet 
                    (this is done to be able to use the validations of changes when the excel is saved)
                */
                if (fe.ReadProperty("conexionSeleccionada", activeSheet.CustomProperties) == null)
                {
                    //If the property is not created,it is established.
                    //In this case a custom property is established for the connection to the table
                    activeSheet.CustomProperties.Add("conexionSeleccionada", JsonConvert.SerializeObject(conexionSeleccionada));
                }
                iniciarDatos(comboBox2.Text);
                this.Close();
            }else
            {
                MessageBox.Show("You must select a table to show the data.");
            }

        }

        public void iniciarDatos(string _tabla)
        {
            //You always work on the active sheet
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            activeSheet.Change -= WorksheetChangeEventHandler;
            //The established connection of the properties of the sheet is obtained
            conexionSeleccionada = JsonConvert.DeserializeObject<conexionesOData>(fe.ReadProperty("conexionSeleccionada", activeSheet.CustomProperties));
            //The OData connection is established
            client = new ODataClient(conexionSeleccionada.Url);
            //Lets create an object to store the connection of the table
            conexionTabla conexiontabla = new conexionTabla();
            conexiontabla.Url = conexionSeleccionada.Url;
            conexiontabla.Tabla = _tabla;

            //The connection property to the table is created for the sheet and the value is assigned
            if (fe.ReadProperty("conexionTabla", activeSheet.CustomProperties) == null)
            {
                activeSheet.CustomProperties.Add("conexionTabla", JsonConvert.SerializeObject(conexiontabla));
            }
            else
            {
                fe.setProperty("conexionTabla", activeSheet.CustomProperties, JsonConvert.SerializeObject(conexiontabla));
            }
            obtenerColumnasDeTabla(activeSheet, _tabla);
            traerDatosDeTabla(conexionSeleccionada.Url,_tabla, true);
        }

        public async void obtenerColumnasDeTabla(Worksheet activeSheet, string _tabla)
        {
            List<string> columnasDeTabla = new List<string>();
            List<string> columnasDeTablaType = new List<string>();
            int idLlavePrimaria = 0;
            int endCol = 0;
            //columnasDeTabla.Clear();
            //columnasDeTablaType.Clear();
            IEdmModel metadata = await client.GetMetadataAsync<IEdmModel>();
            var entityTypes = metadata.SchemaElements.Where(x => x.Name == _tabla).OfType<IEdmEntityType>().ToArray();
            foreach (var tabla in entityTypes)
            {
                var columnas = tabla.DeclaredProperties.ToArray();
                string nombreDeLlavePrimaria = tabla.DeclaredKey.ElementAt(0).Name;
                var cantidad = columnas.Count();
                for (int x = 0; x <= cantidad - 1; x++)
                {
                    var columna = columnas.ElementAt(x);
                    if (columna.PropertyKind.ToString() == "Structural")
                    {
                        columnasDeTablaType.Insert(x, columna.Type.PrimitiveKind().ToString());
                        columnasDeTabla.Insert(x, columna.Name);
                    }

                    if (nombreDeLlavePrimaria == columna.Name)
                    {
                        idLlavePrimaria = x;
                    }
                }

                endCol = columnasDeTabla.Count;
            }
            if (fe.ReadProperty("columnasDeTabla", activeSheet.CustomProperties) == null)
            {
                activeSheet.CustomProperties.Add("columnasDeTabla", JsonConvert.SerializeObject(columnasDeTabla));
                activeSheet.CustomProperties.Add("columnasDeTablaType", JsonConvert.SerializeObject(columnasDeTablaType));
                activeSheet.CustomProperties.Add("idLlavePrimaria", idLlavePrimaria.ToString());
                activeSheet.CustomProperties.Add("endCol", endCol.ToString());
            }
            else
            {
                fe.setProperty("columnasDeTabla", activeSheet.CustomProperties, JsonConvert.SerializeObject(columnasDeTabla));
                fe.setProperty("columnasDeTablaType", activeSheet.CustomProperties, JsonConvert.SerializeObject(columnasDeTablaType));
                fe.setProperty("idLlavePrimaria", activeSheet.CustomProperties, idLlavePrimaria.ToString());
                fe.setProperty("endCol", activeSheet.CustomProperties, endCol.ToString());
            }

        }

        public async void traerDatosDeTabla(string _url,string _tabla, bool mostrarEnExcel)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            List<string> columnasDeTabla = JsonConvert.DeserializeObject<List<string>>(fe.ReadProperty("columnasDeTabla", activeSheet.CustomProperties));
            List<List<string>> listaDeDatosCrudosNew = new List<List<string>>();

            ODataClient client = new ODataClient(_url);
            var informacion = await client.For(_tabla).FindEntriesAsync();
            var datosEnArray = new object[informacion.Count() + 1, columnasDeTabla.Count];
            int indexFila = 1;
            int indexColumna = 1;
            foreach (var col in columnasDeTabla)
            {
                datosEnArray[indexFila - 1, indexColumna - 1] = col;
                indexColumna = indexColumna + 1;
            }

            indexFila = indexFila + 1;
            foreach (var dato in informacion)
            {
                indexColumna = 1;
                List<string> datos_crudos = new List<string>();
                for (int x = 0; x <= columnasDeTabla.Count - 1; x++)
                {
                    var data_ = "";
                    if (dato.ElementAt(x).Value != null)
                    {
                        data_ = dato.ElementAt(x).Value.ToString();
                    }
                    else
                    {
                        data_ = "";
                    }

                    datosEnArray[indexFila - 1, indexColumna - 1] = data_;
                    datos_crudos.Add(data_);
                    indexColumna = indexColumna + 1;
                }
                listaDeDatosCrudosNew.Add(datos_crudos);
                indexFila = indexFila + 1;
            }
            if (fe.ReadProperty("listaDeDatosCrudosNew", activeSheet.CustomProperties) == null)
            {
                activeSheet.CustomProperties.Add("listaDeDatosCrudosNew", JsonConvert.SerializeObject(listaDeDatosCrudosNew));
            }
            else
            {
                fe.setProperty("listaDeDatosCrudosNew", activeSheet.CustomProperties, JsonConvert.SerializeObject(listaDeDatosCrudosNew));
            }
            if (mostrarEnExcel)
            {
                var firstCell = activeSheet.Cells[1, 1];
                var lastCell = activeSheet.Cells[informacion.Count() + 1, columnasDeTabla.Count];

                var range = activeSheet.Range[firstCell, lastCell];
                range.Value2 = datosEnArray;
                range.Columns.AutoFit();
                range.Font.Color = Color.FromArgb(0, 0, 0);

                activeSheet.Change += WorksheetChangeEventHandler;

                var list = activeSheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, range, Type.Missing, XlYesNoGuess.xlYes, Type.Missing);

                list.Name = _tabla;
                list.TableStyle = "TableStyleMedium7";
                
                activeSheet.Name = _tabla;

                Globals.Ribbons.Ribbon1.RibbonUI.ActivateTabMso("TabAddIns");

                var focus = activeSheet.get_Range("A1", "A1");
                focus.Select();
            }
            else
            {
                var firstCell = activeSheet.Cells[1, 1];
                var lastCell = activeSheet.Cells[informacion.Count() + 1, columnasDeTabla.Count];
                var range = activeSheet.Range[firstCell, lastCell];
                range.Font.Color = Color.FromArgb(0, 0, 0);
            }
            
        }

        public void WorksheetChangeEventHandler(Range Target)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            List<List<string>> listaDeDatosCrudosNew = JsonConvert.DeserializeObject<List<List<string>>>(fe.ReadProperty("listaDeDatosCrudosNew", activeSheet.CustomProperties));
            int idLlavePrimaria = Convert.ToInt32(fe.ReadProperty("idLlavePrimaria", activeSheet.CustomProperties));
            try
            {
                int row = Target.Row - 2;
                int col = Target.Column - 1;
                var datosCrudo = listaDeDatosCrudosNew.ElementAt(row);
                var datosNew = Target.Value2;
                bool saber = datosNew is Array;
                if (!saber)
                {

                    int row2 = Target.Row;
                    int columnaid = idLlavePrimaria + 1;
                    var idBuscando = ((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row2, columnaid]).Value2;
                    datosCrudo = listaDeDatosCrudosNew.FindAll(x => x.ElementAt(idLlavePrimaria) == Convert.ToString(idBuscando)).FirstOrDefault();
                    
                    if (datosCrudo.ElementAt(col) != Convert.ToString(datosNew))
                    {
                        Target.Cells.Font.Color = Color.FromArgb(255, 0, 0);
                    }
                    else
                    {
                        Target.Cells.Font.Color = Color.FromArgb(0, 0, 0);
                    }
                }
            }
            catch
            {

            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
