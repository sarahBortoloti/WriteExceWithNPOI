using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel; // para arquivos XLSX

namespace ExcelNPOI
{
    class ArquivoExcelProdutos
    {
        private List<Producto> _produtos;
        private string _nomeArquivo;
        private XSSFWorkbook _workbookCatalogo;

        public ArquivoExcelProdutos(List<Producto> produtos)
        {
            this._produtos = produtos;
        }

        public void GerarArquivo(string nomeArquivo)
        {
            this._nomeArquivo = nomeArquivo;
            _workbookCatalogo = new XSSFWorkbook();
            FileStream file = new FileStream(ConfigurationManager
                .AppSettings["TemplateArquivoExcelProdutos"],
                    FileMode.Open, FileAccess.Read);
            _workbookCatalogo = new XSSFWorkbook(file);

            PreencherInformacoesProdutos();
            FinalizarGravacaoArquivo();
        }

        private void PreencherInformacoesProdutos()
        {   
            ISheet sheetCatalogo = _workbookCatalogo.CreateSheet();

            sheetCatalogo =
                _workbookCatalogo.GetSheet("Hoja1");
            

            int rowIndex = 3;
            foreach (Producto produto in _produtos)
            {

                IRow dataRow = sheetCatalogo.CreateRow(rowIndex);
                
                dataRow.CreateCell(0).SetCellValue(produto.CodigoBarras);
                dataRow.CreateCell(1).SetCellValue(produto.NomeProduto);
                dataRow.CreateCell(2).SetCellValue(produto.CodigoBarras);
                dataRow.CreateCell(3).SetCellValue(produto.Categoria);
                dataRow.CreateCell(4).SetCellValue(produto.QtdEstoque);
                dataRow.CreateCell(5).SetCellValue(produto.PrecoVenda);
            
                rowIndex++;
                /*
                     sheetCatalogo.GetCell(numeroProximaLinha, 1)
                         .SetCellValue(produto.CodigoBarras);
                     sheetCatalogo.GetCell(numeroProximaLinha, 2)
                         .SetCellValue(produto.NomeProduto);
                     sheetCatalogo.GetCell(numeroProximaLinha, 3)
                         .SetCellValue(produto.Categoria);
                     sheetCatalogo.GetCell(numeroProximaLinha, 4)
                         .SetCellValue(produto.DataCadastro);
                     sheetCatalogo.GetCell(numeroProximaLinha, 5)
                         .SetCellValue(produto.QtdEstoque);
                     sheetCatalogo.GetCell(numeroProximaLinha, 6)
                         .SetCellValue(produto.PrecoVenda);
                     numeroProximaLinha++;
                         */

            }

        }

        public void FinalizarGravacaoArquivo()
        {
            using (FileStream file = new FileStream(
                _nomeArquivo, FileMode.Create))
            {
                _workbookCatalogo.Write(file);
                file.Close();
            }
        }
    }
}

