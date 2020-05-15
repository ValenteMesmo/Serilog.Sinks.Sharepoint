using Serilog.Events;
using System;

namespace Serilog
{
    internal class SharepointDefaultFormatter : ICustomFormatter
    {
        public string Format(string format, object arg, IFormatProvider formatProvider)
        {
            if (!(arg is LogEvent logEvent))
                return format;

            var sufix = logEvent.Exception == null
                ? ""
                : logEvent.Exception.ToString();

            if (string.IsNullOrEmpty(sufix))
                return format;

            return $@"{format}

=========================================
{sufix}";
        }
    }
}
