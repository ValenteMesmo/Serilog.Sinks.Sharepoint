using Microsoft.SharePoint.Client;
using Serilog.Core;
using Serilog.Events;
using System;

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

            var createView = false;
            if (!context.ListExists(listName).GetAwaiter().GetResult())
            {
                context.CreateList(listName).GetAwaiter().GetResult();
                createView = true;
            }

            if (!list.ContainsField(detailFieldName).GetAwaiter().GetResult())
                list.AddNoteField(detailFieldName).GetAwaiter().GetResult();

            if (createView)
            {
                var view = list.DefaultView;
                view.ViewFields.RemoveAll();
                view.ViewFields.Add("Created");
                view.ViewFields.Add("Title");
                view.ViewFields.Add(detailFieldName);
                view.ViewFields.Add("Author");
                view.ViewQuery = "<OrderBy><FieldRef Name='ID' Ascending='FALSE' /></OrderBy>";
                view.Update();
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
}
