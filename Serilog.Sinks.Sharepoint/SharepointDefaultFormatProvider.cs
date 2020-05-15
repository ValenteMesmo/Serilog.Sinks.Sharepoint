using System;

namespace Serilog
{
    internal class SharepointDefaultFormatProvider : IFormatProvider
    {
        public object GetFormat(Type formatType)
        {
            return new SharepointDefaultFormatter();
        }
    }
}
