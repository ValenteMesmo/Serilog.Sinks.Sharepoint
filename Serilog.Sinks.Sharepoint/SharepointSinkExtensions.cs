using Microsoft.SharePoint.Client;
using Serilog.Configuration;
using System;

namespace Serilog
{
    public static class SharepointSinkExtensions
    {
        public static LoggerConfiguration Sharepoint(
                  this LoggerSinkConfiguration loggerConfiguration,
                  ClientContext context,
                  string listName = "Log",
                  string detailFieldName = "LogDetail",
                  IFormatProvider formatProvider = null)
        {
            return loggerConfiguration.Sink(new SharepointSink(context, listName, detailFieldName, formatProvider));
        }
    }
}
