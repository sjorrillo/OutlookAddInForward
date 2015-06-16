
namespace OutlookAddInForward.Extensions
{
    public class TemplateResult
    {
        public TemplateResult(string template, bool success = true)
        {
            Success = success;
            CompiledTemplate = template;
        }

        public bool Success { get; private set; }

        public string CompiledTemplate { get; private set; }
    }
}
