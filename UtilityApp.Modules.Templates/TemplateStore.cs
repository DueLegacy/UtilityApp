using System.Collections.Generic;
using System.Xml.Serialization;

namespace UtilityApp.Modules.Templates
{
    [XmlRoot("TemplateStore")]
    public class TemplateStore
    {
        public TemplateStore()
        {
            Templates = new List<TemplateItem>();
        }

        public TemplateStore(IEnumerable<TemplateItem> templates)
            : this()
        {
            if (templates == null)
            {
                return;
            }

            foreach (var template in templates)
            {
                Templates.Add(template);
            }
        }

        [XmlElement("Template")]
        public List<TemplateItem> Templates { get; set; }
    }
}
