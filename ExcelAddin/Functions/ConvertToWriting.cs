using System.Text;
using ExcelDna.Integration;
using GeoGroup.ExcelExtension.External;

namespace GeoGroup.ExcelExtension.Functions
{
    public class ConvertToWriting
    {
        [ExcelFunction(
            Name = "Прописью",
            Description = "Функция принимает число и возвращает результат прописью", 
            Category = "ВитГеоГрупп")]
        public static string ConvertToWritingFunction(
            [ExcelArgument(
                Name = "Параметр", 
                Description = "Число, которое будет преобразовано в пропись")]
            decimal valueToConvert)
        {
            StringBuilder resultBuilder = new StringBuilder();
            WritingSumConverter.ToWriting(valueToConvert, Currency.Рубли, resultBuilder);
            return resultBuilder.ToString();
        }
    }
}
