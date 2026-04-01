using Umbraco.Cms.Core.Composing;
using Umbraco.Cms.Core.DependencyInjection;

namespace Umbraco.Community.Examine.OpenXml
{
    /// <summary>
    /// Registers the ExamineOpenXml index, and dependencies.
    /// </summary>
    public class ExamineOpenXmlComposer : IComposer
    {
        public void Compose(IUmbracoBuilder builder) => builder.AddExamineOpenXml();
    }
}