using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Excel_generalas
{
    public partial class Form1 : Form
    {
        RealEstateEntities context = new RealEstateEntities(); // ORM objektum példányosítása
        List<Flat> Flats; // Flat tipusú elemekből álló listára mutató referencia
        public Form1()
        {
            InitializeComponent();
            LoadData();
        }

        private void LoadData()
        {
            Flats = context.Flats.ToList(); //Adattábla memóriába való másolása
        }
    }
}
