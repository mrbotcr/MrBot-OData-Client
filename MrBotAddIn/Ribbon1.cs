using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using MrBotAddIn.Properties;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Simple.OData.Client;
using System.Drawing;
using System.Threading.Tasks;

namespace MrBotAddIn
{
    public partial class Ribbon1
    {
        funcionesEspeciales fe = new funcionesEspeciales();
        public List<conexionesOData> listaDeConexiones = new List<conexionesOData>();
        public Resultados formResultados = new Resultados();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //formResultados = new Resultados();
        }
        
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            connections connections = new connections();
            connections.Show();            
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                guardarDatosDeTabla();
            }catch
            {
                MessageBox.Show("Connection not established.");
            }
        }

        public void guardarDatosDeTabla()
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            int row = 2;
            int col = 1;
            List<List<string>> listaDeDatosCrudosNew = JsonConvert.DeserializeObject<List<List<string>>>(fe.ReadProperty("listaDeDatosCrudosNew", activeSheet.CustomProperties));
            List<List<string>> listaDeDatosActualizados = new List<List<string>>();
            List<string> llavesPrimarias = new List<string>();
            int endCol = Convert.ToInt32(fe.ReadProperty("endCol", activeSheet.CustomProperties));
            int index  = Convert.ToInt32(fe.ReadProperty("idLlavePrimaria", activeSheet.CustomProperties));

            while (((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row, col]).Value2 != null)
            {
                List<string> datos_act = new List<string>();

                for (int y = 1; y <= endCol; y++)
                {
                    var data_ = ((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row, col]).Value2;
                    if (data_ is double)
                    {
                        data_ = ((double)((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row, col]).Value2).ToString();
                    }

                    if (data_ == null)
                    {
                        datos_act.Add("");
                    }
                    else
                    {
                        datos_act.Add(data_);
                    }

                    if (index == y - 1)
                    {
                        llavesPrimarias.Add(data_);
                    }
                    col++;
                }
                row++;
                col = 1;
                listaDeDatosActualizados.Add(datos_act);

            }

            if (llavesPrimarias.Count.ToString() == llavesPrimarias.Distinct().Count().ToString())
            {
                revisarInformacion(listaDeDatosActualizados, listaDeDatosCrudosNew);

            }
            else
            {
                MessageBox.Show("No se puede actualizar por que hay llaves repetidas.");
            }
        }

        public void revisarInformacion(List<List<string>> listaDeDatosActualizados, List<List<string>> listaDeDatosCrudos)
        {
            formResultados.Show();
            formResultados.progressBar1.Maximum = listaDeDatosActualizados.Count;
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            conexionTabla conexiontabla = JsonConvert.DeserializeObject<conexionTabla>(fe.ReadProperty("conexionTabla", activeSheet.CustomProperties));
            ODataClient client = new ODataClient(conexiontabla.Url);
            int index = Convert.ToInt32(fe.ReadProperty("idLlavePrimaria", activeSheet.CustomProperties));
            int contador = 2;
            foreach (List<string> dato_actualizado in listaDeDatosActualizados)
            {
                List<List<string>> listaExistentes = listaDeDatosCrudos.FindAll(x => x.ElementAt(index) == dato_actualizado[index]);
                if (listaExistentes.Count > 0)
                {
                    foreach (List<string> trabajando in listaExistentes)
                    {
                        if (trabajando.SequenceEqual(dato_actualizado))
                        {
                            //MessageBox.Show("No cambio: " + dato_actualizado[1]);
                        }
                        else
                        {
                            //MessageBox.Show("Cambio: " + dato_actualizado[1]);
                            IDictionary<string, object> dictionary = crearDiccionarioUpdate(dato_actualizado, trabajando);
                            var ship = client.For(conexiontabla.Tabla).Key(dictionary).Set(dictionary).UpdateEntryAsync();
                            
                            try
                            {
                                ship.Wait();
                                formResultados.richTextBox1.SelectionColor = Color.Green;
                                formResultados.richTextBox1.AppendText("Línea "+contador.ToString()+ " - UPDATE - Correcto.\n");
                                //MessageBox.Show("Bien actualizado.\n");
                            }
                            catch
                            {
                                formResultados.richTextBox1.SelectionColor = Color.Red;
                                formResultados.richTextBox1.AppendText("Línea "+contador.ToString()+ " - UPDATE - Error.\n");
                                //MessageBox.Show(ship.Exception.Message);
                                //MessageBox.Show(ship.Status.ToString());
                            }
                        }
                        listaDeDatosCrudos.Remove(trabajando);
                    }
                }
                else
                {
                    //MessageBox.Show("Insertar: " + dato_actualizado[1]);
                    IDictionary<string, object> dictionary = crearDiccionario(dato_actualizado);
                    var ship = client.For(conexiontabla.Tabla).Set(dictionary).InsertEntryAsync();
                    try
                    {
                        ship.Wait();
                        //MessageBox.Show("Dato agregado correctamente.");
                        formResultados.richTextBox1.SelectionColor = Color.Green;
                        formResultados.richTextBox1.AppendText("Línea " + contador.ToString() + " - INSERT - Correcto.\n");
                    }
                    catch
                    {
                        //MessageBox.Show(ship.Exception.Message);
                        //MessageBox.Show(ship.Status.ToString());
                        formResultados.richTextBox1.SelectionColor = Color.Red;
                        formResultados.richTextBox1.AppendText("Línea " + contador.ToString() + " - INSERT - Error.\n");
                    }
                }
                contador = contador + 1;
                formResultados.progressBar1.Value = formResultados.progressBar1.Value + 1;
            }

            foreach (List<string> eliminar in listaDeDatosCrudos)
            {
                //MessageBox.Show("Eliminar: " + eliminar[1]);
                IDictionary<string, object> dictionary = crearDiccionario(eliminar);
                var ship = client.For(conexiontabla.Tabla).Key(dictionary).DeleteEntriesAsync();
                try
                {
                    ship.Wait();
                    //MessageBox.Show("Dato eliminado");
                    formResultados.richTextBox1.SelectionColor = Color.Green;
                    formResultados.richTextBox1.AppendText("Línea - DELETE - Correcto.\n");
                }
                catch
                {
                    //MessageBox.Show(ship.Exception.Message);
                    //MessageBox.Show(ship.Status.ToString());
                    formResultados.richTextBox1.SelectionColor = Color.Red;
                    formResultados.richTextBox1.AppendText("Línea - DELETE - Error.\n");
                }
            }
            formResultados = new Resultados();
            connections ri = new connections();
            ri.traerDatosDeTabla(conexiontabla.Url, conexiontabla.Tabla,false);
        }

        public IDictionary<string, object> crearDiccionario(List<string> dato_actualizadoinput)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            List<string> dato_actualizado = dato_actualizadoinput;
            IDictionary<string, object> dictionary = new Dictionary<string, object>();
            List<string> columnasDeTabla = JsonConvert.DeserializeObject<List<string>>(fe.ReadProperty("columnasDeTabla", activeSheet.CustomProperties));
            List<string> columnasDeTablaType = JsonConvert.DeserializeObject<List<string>>(fe.ReadProperty("columnasDeTablaType", activeSheet.CustomProperties)); ;

            int cantidadDeColumnas = columnasDeTabla.Count;
            for (int vali = 0; vali <= cantidadDeColumnas - 1; vali++)
            {
                dictionary.Add(columnasDeTabla.ElementAt(vali), dato_actualizado.ElementAt(vali));
            }

            return dictionary;
        }

        public IDictionary<string, object> crearDiccionarioUpdate(List<string> dato_actualizadoinput, List<string> trabajandoinput)
        {
            List<string> dato_actualizado = dato_actualizadoinput;
            IDictionary<string, object> dictionary = new Dictionary<string, object>();
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            List<string> columnasDeTabla = JsonConvert.DeserializeObject<List<string>>(fe.ReadProperty("columnasDeTabla", activeSheet.CustomProperties));
            List<string> columnasDeTablaType = JsonConvert.DeserializeObject<List<string>>(fe.ReadProperty("columnasDeTablaType", activeSheet.CustomProperties)); ;
            int idLlavePrimaria = Convert.ToInt32(fe.ReadProperty("idLlavePrimaria", activeSheet.CustomProperties));
            int cantidadDeColumnas = columnasDeTabla.Count;
            for (int vali = 0; vali <= cantidadDeColumnas - 1; vali++)
            {
                if (dato_actualizado.ElementAt(vali) != trabajandoinput.ElementAt(vali) || vali == idLlavePrimaria)
                {
                    var saberType = columnasDeTablaType.ElementAt(vali);
                    if (saberType == "DateTime")
                    {
                        double date = double.Parse(dato_actualizado.ElementAt(vali).ToString());
                        var fecha = DateTime.FromOADate(date).ToString("o");
                        dictionary.Add(columnasDeTabla.ElementAt(vali), fecha);
                    }
                    else
                    {
                        dictionary.Add(columnasDeTabla.ElementAt(vali), dato_actualizado.ElementAt(vali));
                    }
                }
            }

            return dictionary;
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
                conexionTabla conexiontabla = JsonConvert.DeserializeObject<conexionTabla>(fe.ReadProperty("conexionTabla", activeSheet.CustomProperties));
                connections ri = new connections();
                ri.iniciarDatos(conexiontabla.Tabla);
            }catch
            {
                MessageBox.Show("Problem loading the data.");
            }
        }

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            this.RibbonUI.ActivateTab(this.Tabs[1].ControlId.ToString());
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            about ab = new about();
            ab.Show();
        }
    }
}
