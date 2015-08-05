using MarkupConverter;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

namespace MarkupConverter
{


    public interface IMarkupConverter
    {
        string ConvertXamlToHtml(string xamlText);
        string ConvertHtmlToXaml(string htmlText);
        string ConvertRtfToHtml(string rtfText);
        string ConvertHtmlToRtf(string htmlText);
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public class MarkupConverter : IMarkupConverter
    {
        public string Text
        {
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
            [param: MarshalAs(UnmanagedType.BStr)]
            set;
        }

        [return: MarshalAs(UnmanagedType.BStr)]
        public string ConvertXamlToHtml(string xamlText)
        {
            return HtmlFromXamlConverter.ConvertXamlToHtml(xamlText, false);
        }

        [return: MarshalAs(UnmanagedType.BStr)]
        public string ConvertHtmlToXaml(string htmlText)
        {
            return HtmlToXamlConverter.ConvertHtmlToXaml(htmlText, true);
        }

        [return: MarshalAs(UnmanagedType.BStr)]
        public string ConvertRtfToHtml(string rtfText)
        {
            return RtfToHtmlConverter.ConvertRtfToHtml(rtfText);
        }

        [return: MarshalAs(UnmanagedType.BStr)]
        static public string ConvertRtfToHtmlEasy(string rtfText)
        {
            return RtfToHtmlConverter.ConvertRtfToHtml(rtfText);
        }

        [return: MarshalAs(UnmanagedType.BStr)]
        public string ConvertHtmlToRtf(string htmlText)
        {
            return HtmlToRtfConverter.ConvertHtmlToRtf(htmlText);
        }

        [return: MarshalAs(UnmanagedType.IDispatch)]
        static object CreateDotNetObject(string text)
        {
            return new MarkupConverter { Text = text };
        }
    }
    /*
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    static class UnmanagedExports
    {
        //[DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        static object CreateDotNetObject(string text)
        {
            return new MarkupConverter { Text = text };
        }
    }
     */
}
