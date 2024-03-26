using ImportarPlanilhaExcel.Data;
using ImportarPlanilhaExcel.Models;
using OfficeOpenXml;

namespace ImportarPlanilhaExcel.Services
{
    public class ExcelService : IExcelInterface
    {
        private readonly AppDbContext _context;
        public ExcelService(AppDbContext context) 
        {
            _context = context;
        }
        public MemoryStream LerStream(IFormFile formfile)
        {
            using (var stream = new MemoryStream())
            {
                formfile?.CopyTo(stream);
                var listaBytes = stream.ToArray();
                return new MemoryStream(listaBytes);
            }
        }

        public List<ProdutoModel> LerXls(MemoryStream stream)
        {
            try
            {
                var resposta = new List<ProdutoModel>();

                ExcelPackage.LicenseContext =  LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(stream)) 
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int numeroLinhas = worksheet.Dimension.End.Row;

                    for(int linha=2; linha <= numeroLinhas; linha++) 
                    {
                        var produto = new ProdutoModel();

                        if (worksheet.Cells[linha, 2].Value != null && worksheet.Cells[linha, 5].Value != null) 
                        {
                            produto.Id = 0;
                            produto.Nome = worksheet.Cells[linha, 2].Value.ToString();
                            produto.Valor = Convert.ToDecimal(worksheet.Cells[linha, 3].Value);
                            produto.Quantidade = Convert.ToInt32(worksheet.Cells[linha, 4].Value);
                            produto.Marca = worksheet.Cells[linha, 5].Value.ToString();

                            resposta.Add(produto);
                        }
                    }
                }
                return resposta;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void SalvarDados(List<ProdutoModel> produtos)
        {
            try
            {
                //foreach(var produto in produtos) 
                //{
                //    _context.Add(produto);
                //    _context.SaveChangesAsync();
                //}
                _context.Produtos.RemoveRange(_context.Produtos);
                _context.SaveChanges();

                _context.Produtos.AddRange(produtos);
                _context.SaveChanges();
            }
            catch (Exception ex) 
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
