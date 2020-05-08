using Microsoft.SharePoint.Client;
using Serilog.Configuration;
using Serilog.Core;
using Serilog.Events;
using System;
using System.Linq;

namespace Serilog
{
    public class SharepointSink : ILogEventSink
    {
        private readonly IFormatProvider FormatProvider;
        private readonly ClientContext context;
        private readonly string detailFieldName;
        private readonly List list;

        public SharepointSink(
            ClientContext context
            , string listName
            , string detailFieldName
            , IFormatProvider FormatProvider = null
        )
        {
            this.context = context;
            this.detailFieldName = detailFieldName;
            ListCollection listCollection = context.Web.Lists;

            context.Load(listCollection, lists => lists
                .Include(list => list.Title)
                .Where(list => list.Title == listName));

            context.ExecuteQuery();

            if (listCollection.Count > 0)
            {
                this.list = context.Web.Lists.GetByTitle(listName);
            }
            else
            {
                this.list = context.Web.Lists.Add(
                    new ListCreationInformation()
                    {
                        Title = listName,
                        TemplateType = (int)ListTemplateType.GenericList
                    });
                context.ExecuteQuery();
            }

            context.Load(list.Fields);
            context.ExecuteQuery();

            if (!list.Fields.Any(f => f.InternalName == detailFieldName))
            {
                list.Fields.AddFieldAsXml(
                    $"<Field Type='Note' DisplayName='{detailFieldName}'/>", true, AddFieldOptions.AddFieldToDefaultView);

                list.Update();
                context.ExecuteQuery();
            }

            this.FormatProvider = FormatProvider ?? new SharepointDefaultFormatProvider();
        }

        public void Emit(LogEvent logEvent)
        {
            try
            {
                var item = list.AddItem(new ListItemCreationInformation() { });

                var message = logEvent.RenderMessage();

                message = (FormatProvider.GetFormat(typeof(LogEvent)) as ICustomFormatter)
                    .Format(message, logEvent, null);

                item["Title"] = logEvent.Level.ToString();
                item[detailFieldName] = message;
                item.Update();
                context.ExecuteQuery();
            }
            catch
            {
            }
        }
    }

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

    internal class SharepointDefaultFormatProvider : IFormatProvider
    {
        public object GetFormat(Type formatType)
        {
            return new SharepointDefaultFormatter();
        }
    }
}
