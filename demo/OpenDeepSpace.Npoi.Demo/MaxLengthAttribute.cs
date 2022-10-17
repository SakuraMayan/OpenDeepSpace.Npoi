using OpenDeepSpace.Npoi.Attributes;

namespace OpenDeepSpace.Npoi.Demo
{
    public class MaxLengthAttribute : DataValidationAttribute
    {

        public int MaxLength { get; set; }

        public override DataValidationResult IsValid(object data)
        {
            if (data.ToString().Length > MaxLength)
                return new DataValidationResult(ErrorMessage);

            return DataValidationResult.Success;
        }
    }
}
