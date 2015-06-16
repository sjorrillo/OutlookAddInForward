
namespace OutlookAddInForward.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Xml;
    using System.Xml.Linq;
    using System.Xml.XPath;
    using System.Xml.Xsl;

    public static class EmailTemplate
    {
        #region Xsl Templates

        /// <summary>
        /// Generates the E-mail body content using a XSL Transformation and a EmailTemplateParameters object 
        /// to set the input data required by the XSL Template.
        /// </summary>
        /// <param name="templatePath">XSL Tranformation template</param>
        /// <param name="templateProperties">Properties required by the XSL Transformation template</param>
        /// <returns>Instance of TemplateResult generated</returns>
        public static TemplateResult GenerateEmailBodyFromXslFile(string templatePath, Dictionary<string, object> templateProperties)
        {
            var document = new XDocument(new XDeclaration("1.0", "utf-8", "yes"),
                new XElement("properties",
                    from record in templateProperties
                    select new XElement(record.Key, record.Value)));

            return GenerateEmailBodyFromXslFile(templatePath, document);
        }

        /// <summary>
        /// Generates the E-mail body content using a XSL Transformation and a XNode object 
        /// with the input data required by the XSL Template.
        /// </summary>
        /// <param name="templatePath">XSL Tranformation template</param>
        /// <param name="xmlInputData">XML input data required by the XSL Transformation template</param>
        /// <returns>Instance of TemplateResult generated</returns>
        public static TemplateResult GenerateEmailBodyFromXslFile(string templatePath, XNode xmlInputData)
        {
            var success = false;
            var compiledTemplate = string.Empty;
            var memoryStream = new MemoryStream();
            var streamWriter = new StreamWriter(memoryStream, System.Text.Encoding.UTF8);
            var htmlWriter = new XmlTextWriter(streamWriter);
            var streamReader = new StreamReader(memoryStream);
            try
            {
                if (templatePath.StartsWith("~"))
                {
                    var baseDir = System.AppDomain.CurrentDomain.BaseDirectory;
                    templatePath = Path.GetFullPath(baseDir + templatePath.Replace("~", ""));
                }
                var xslCompiledTransform = new XslCompiledTransform();
                xslCompiledTransform.Load(templatePath);

                xslCompiledTransform.Transform(xmlInputData.CreateNavigator(), htmlWriter);

                memoryStream.Position = 0;
                compiledTemplate = streamReader.ReadToEnd();
                success = true;
            }
            catch (XsltException xsltException)
            {
                compiledTemplate = string.Format(CultureInfo.InvariantCulture, "Error: {0}\n\tFileName: {1}\n\tLine Number: {2} - Position: {3}", new object[] { xsltException.Message, xsltException.SourceUri, xsltException.LineNumber, xsltException.LinePosition });
            }
            catch (Exception ex)
            {
                compiledTemplate = string.Format(CultureInfo.InvariantCulture, "Error: {0}", ex.Message);
            }
            finally
            {
                htmlWriter.Close();
                streamReader.Close();
                memoryStream.Close();
            }
            return new TemplateResult(compiledTemplate, success);
        } 

        #endregion
    }
}
