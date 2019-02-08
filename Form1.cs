using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelNPOI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnGerarCatalogo_Click(
          object sender, EventArgs e)
        {
            if (saveFileDialogCatalogo.ShowDialog() ==
                DialogResult.OK)
            {
                ArquivoExcelProdutos arq = new ArquivoExcelProdutos(
                    CatalogoProductoscs.ObterCatalogo());
                arq.GerarArquivo(saveFileDialogCatalogo.FileName);
                MessageBox.Show("O arquivo " +
                    saveFileDialogCatalogo.FileName +
                    " foi gerado com sucesso!");
            }
        }

     
    }
}


