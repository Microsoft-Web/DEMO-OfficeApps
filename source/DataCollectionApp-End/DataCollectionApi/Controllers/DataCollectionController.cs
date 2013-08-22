using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace DataCollectionApi.Controllers
{
    public class DataCollectionController : ApiController
    {
        public string Post(JToken parm)
        {
            return string.Format("{0} {1} saved", parm["first"], parm["last"]);
        }
    }
}
