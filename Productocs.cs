using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


    namespace ExcelNPOI
    {
        public class Producto
        {
            public string CodigoBarras { get; set; }
            public string NomeProduto { get; set; }
            public string Categoria { get; set; }
            public DateTime DataCadastro { get; set; }
            public int QtdEstoque { get; set; }
            public double PrecoVenda { get; set; }
        }
    }

