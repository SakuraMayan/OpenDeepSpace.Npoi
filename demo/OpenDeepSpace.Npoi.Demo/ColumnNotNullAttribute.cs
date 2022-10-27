using OpenDeepSpace.Npoi.Attributes;

namespace OpenDeepSpace.Npoi.Demo
{
    public class ColumnNotNullAttribute : DataValidationAttribute
    {
        public override DataValidationResult? IsValid(object data)
        {
            if (data == null || string.IsNullOrWhiteSpace(data.ToString()))
                return new DataValidationResult(ErrorMessage);
            return DataValidationResult.Success;
        }

    }
}
