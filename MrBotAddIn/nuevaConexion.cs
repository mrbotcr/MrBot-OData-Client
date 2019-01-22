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
using System.Runtime.Serialization.Json;
using System.IO;
using System.Collections.Specialized;
using Newtonsoft.Json;
using Microsoft.Data.Edm;
using System.Net;

namespace MrBotAddIn
{
    public partial class nuevaConexion : Form
    {
        public Ribbon1 ribbon = new Ribbon1();
        
        conexionesOData datosDeConexion = new conexionesOData();

        public nuevaConexion()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            propertyGrid1.SelectedObject = datosDeConexion;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (datosDeConexion.Name != "")
                {
                    /*
                        In the properties of the complement we are storing the connections, 
                        then we obtain the existing connections to add the new one.
                        If we do not have a connection we establish that we will create a list of connections.
                        We validate that there is no connection with a duplicate name and 
                        add it to the list of connections, write it in the properties of the project and save the information.
                    */
                    List<conexionesOData> lista = JsonConvert.DeserializeObject<List<conexionesOData>>(MrBotAddIn.Properties.Settings.Default.jsonDeConexiones);
                    if (lista == null)
                        lista = new List<conexionesOData>();
                    if (lista.Where(x => x.Name == datosDeConexion.Name).Count() == 0)
                    {
                        lista.Add(datosDeConexion);
                        MrBotAddIn.Properties.Settings.Default.jsonDeConexiones = JsonConvert.SerializeObject(lista);
                        MrBotAddIn.Properties.Settings.Default.Save();
                        
                        connections connections = new connections();
                        connections.Show();
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("You already have a connection with this name.");
                    }
                }
                else
                {
                    MessageBox.Show("Enter a new name for the connection.");
                }
            }catch
            {
                MessageBox.Show("The connection could not be created, check the provided URL.");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            connections connections = new connections();
            connections.Show();
            this.Close();
        }

        private void propertyGrid1_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            button2.Enabled = false;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            try
            {
                /* 
                    We create an Odata Client Settings to define the URL 
                    with which we establish the connection
                */
                ODataClientSettings odcSettings = new ODataClientSettings();
                //Define the URL
                Uri uriOdata = new Uri(datosDeConexion.Url, UriKind.Absolute);
                odcSettings.BaseUri = uriOdata;
                /*
                We establish the connection and we bring the metadata to know if it is connected
                */
                ODataClient client = new ODataClient(odcSettings);
                IEdmModel metadata = await client.GetMetadataAsync<IEdmModel>();
                var entityTypes = metadata.SchemaElements.OfType<IEdmEntityType>().ToArray();
                button2.Enabled = true;
            }catch
            {
                MessageBox.Show("Error creating the connection.");
            }
        }
    }

}
