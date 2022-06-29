using System.Threading.Tasks;
using TestReadFile.Models;

namespace TestReadFile.Interfaces
{
    public interface IWriter
    {
        void Write(Sheet1[] rows, string path);
    }
}
