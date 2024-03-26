using ImportarPlanilhaExcel.Models;

namespace ImportarPlanilhaExcel.Services
{
    public interface IExcelInterface
    {
        MemoryStream LerStream(IFormFile formfile);
        List<ProdutoModel> LerXls(MemoryStream stream);
        void SalvarDados(List<ProdutoModel> produtos);
    }
}
