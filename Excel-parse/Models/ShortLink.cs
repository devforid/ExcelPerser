using System;
using System.Collections.Generic;
using System.Text;

namespace Excel_parse.Models
{
    public class ShortLink
    {
        public string RequestUrl { get; set; }
        public string RequestMethod { get; set; }
        public string RequestPayload { get; set; }
        public string RequestEncodedQueryString { get; set; }
        public string RequestByUserId { get; set; }
        public string RequestHeaders { get; set; }
        public string RedirectUrl { get; set; }
        public string ExpiryLifeSpan { get; set; }
        public string UseRequestLimit { get; set; }
        public string RequestLimit { get; set; }
        public string UserCanLogin { get; set; }
        public string LinkBasedActionConfigId { get; set; }

    }
}
