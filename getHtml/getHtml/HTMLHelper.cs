using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Text.RegularExpressions;
using System.IO;
using System.Net.Security;


namespace getHtml
{

    /// <summary>
    /// html代码处理基类
    /// </summary>
    public static class HTMLHelper
    {


        /// <summary>
        /// 获取某个网站页面的html代码
        /// </summary>
        /// <param name="url">网站路径</param>
        /// <returns></returns>
        public static string GetHTMLByUrl(string url, string encoding)
        {
            HttpWebRequest myHttpWebRequest = null;
            try
            {
                string strResult = "";
                //如果是发送HTTPS请求
                if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
                {
                    ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;
                    myHttpWebRequest = WebRequest.Create(url) as HttpWebRequest;
                }
                else
                {
                    myHttpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                }
                myHttpWebRequest.ContentType = "text/html";
                myHttpWebRequest.Method = "GET";
                myHttpWebRequest.Referer = url;
                myHttpWebRequest.UserAgent = "Mozilla/5.0 (Windows; U; Windows NT 5.2; zh-CN; rv:1.9.1.10) Gecko/20100504 Firefox/3.5.10";
                myHttpWebRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                HttpWebResponse response = (HttpWebResponse)myHttpWebRequest.GetResponse();
                Stream stream = response.GetResponseStream();
                stream.ReadTimeout = 10000;
                System.IO.StreamReader sr = new System.IO.StreamReader(stream, Encoding.GetEncoding(encoding));     //utf-8

                strResult = sr.ReadToEnd();
                sr.Close();
                stream.Close();
                response.Close();
                return strResult;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            {
                if (myHttpWebRequest != null)
                {
                    myHttpWebRequest.Abort();
                }
            }
        }

        /// <summary>
        /// 获取某个网站页面的html代码
        /// </summary>
        /// <param name="url">网站路径</param>
        /// <returns></returns>
        public static string GetHTMLByUrl(string url, string encoding, string setCookie)
        {
            HttpWebRequest myHttpWebRequest = null;
            try
            {
                string strResult = "";
                //如果是发送HTTPS请求
                if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
                {
                    ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;
                    myHttpWebRequest = WebRequest.Create(url) as HttpWebRequest;
                }
                else
                {
                    myHttpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                }
                myHttpWebRequest.ProtocolVersion = HttpVersion.Version10;
                myHttpWebRequest.ContentType = "application/x-www-form-urlencoded";
                myHttpWebRequest.Method = "GET";
                myHttpWebRequest.KeepAlive = false;
                myHttpWebRequest.Timeout = 60000;
                //myHttpWebRequest.ContentType = "application/x-www-form-urlencoded";
                myHttpWebRequest.Proxy = null;
                myHttpWebRequest.ServicePoint.Expect100Continue = false;
                //是否使用 Nagle 不使用 提高效率 
                myHttpWebRequest.ServicePoint.UseNagleAlgorithm = false;
                //最大连接数 
                myHttpWebRequest.ServicePoint.ConnectionLimit = 65500;
                //数据是否缓冲 false 提高效率  
                myHttpWebRequest.AllowWriteStreamBuffering = false;

                
                myHttpWebRequest.AllowAutoRedirect = true;

                myHttpWebRequest.UserAgent = "Mozilla/5.0 (iPhone; CPU iPhone OS 7_0_4 like Mac OS X) AppleWebKit/537.51.1 (KHTML, like Gecko) Mobile/11B554a MicroMessenger/5.1";// "Mozilla/5.0 (Windows NT 6.1; rv:23.0) Gecko/20100101 Firefox/23.0";
                myHttpWebRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                myHttpWebRequest.Headers.Add("Accept-Language", "zh-cn,en-us;q=0.5");

                myHttpWebRequest.Referer = url;
                //myHttpWebRequest.Connection = "keep-alive";
               // setCookie = "	Sun, 16 Feb 2014 08:10:23 GMT; pt2gguin=o1053352248; ptcz=9dc77de6369ca243da579b76f84776bf522f7acf4a033809635741ae1bb219c1; pgv_pvid=9628362689; o_cookie=1053352248; ptisp=ctc; qm_sid=0b457f5527c24ec4821da769184afec9,qN2o2a0ZLQUswbWFFbnQqZE5hSThpMFBRVi0xRipGbHJFbDhlcVdBcXlHZ18.; qm_username=1053352248; pgv_info=ssid=s9922989580; uin=o1053352248; skey=@jmKlUBH4z; RK=pb2+JyxBcs; slave_user=gh_8c17aa9dc465; slave_sid=RE40NnMxQzVIczZLUFNhUXVYYmp1XzRpOGZiaHlqczk2d2R3eFBDZGVHVzJ2M1FaWTFnc1prSTR6VGw1Yk5lZmdwTU5jbFpxMFNydlRnclJmSGllMFlneVFDRFptUGVCYl9sbmloT2k4ZUwxdWJRSVZTVTgwK3hSbFkyRWprbmk=; bizuin=2391051669";
                if (setCookie != null)
                    myHttpWebRequest.Headers["Cookie"] = setCookie;



                HttpWebResponse response = (HttpWebResponse)myHttpWebRequest.GetResponse();
                Stream stream = response.GetResponseStream();
                stream.ReadTimeout = 10000;
                System.IO.StreamReader sr = new System.IO.StreamReader(stream, Encoding.GetEncoding(encoding));     //utf-8
               
                strResult = sr.ReadToEnd();
                sr.Close();
                stream.Close();
                response.Close();
                return strResult;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            {
                if (myHttpWebRequest != null)
                {
                    myHttpWebRequest.Abort();
                }
            }
        }

        /// <summary>
        /// 处理文章列表区域、文章标题、文章内容
        /// 返回指定规则的HTML内容
        /// </summary>
        /// <param name="strHTML">源代码</param>
        /// <param name="rule">规则</param>
        /// <returns></returns>
        public static string GetHtmlContent(string strHTML, string rule)
        {
            if (strHTML == null || strHTML == "" || rule == null || rule == "")
            {
                return "";
            }
            Regex r = new Regex(rule, RegexOptions.IgnoreCase);
            Match m = r.Match(strHTML);
            if (m.Success)
            {
                GroupCollection gc = m.Groups;
                string value;
                if (gc != null && gc.Count > 1)
                {
                    for(int i=1;i<gc.Count;i++){
                        value = gc[i].Value.Trim();
                        if (value.Length > 0)
                            return value;
                    }
                   // return gc[1].Value;
                }

            }
            return "";
        }

        /// <summary>
        /// 返回匹配的字符串集合
        /// </summary>
        /// <param name="strHTML">源代码</param>
        /// <param name="rule">规则</param>
        /// <returns></returns>
        public static List<string> GetMatchString(string strHTML, string rule)
        {
            // rule = "<dd>((?:(?!id=1766690474)[\\d\\D])+?)dd>";
            if (strHTML == null || strHTML == "" || rule == null || rule == "")
            {
                return null;
            }
            try
            {
                Regex r = new Regex(rule, RegexOptions.IgnoreCase);
                MatchCollection mc = r.Matches(strHTML);
                if (mc != null && mc.Count > 0)
                {
                    List<string> list = new List<string>();
                    string value;
                    foreach (Match m in mc)
                    {
                        GroupCollection gc = m.Groups;
                        if (gc != null && gc.Count > 1)
                        {
                            for (int i = 1; i < gc.Count; i++)
                            {
                                value = gc[i].Value.Trim();
                                if (value.Length > 0)
                                {
                                    list.Add(value);
                                }
                            }
                        }
                    }
                    return list;
                }
            }
            catch (Exception) { }

            return null;
        }

        

       

        /// <summary>
        /// 创建正则模型
        /// </summary>
        /// <param name="rule"></param>
        /// <returns></returns>
        private static string CreatePattern(string rule)
        {
            //  return rule;
            rule = rule.ToLower().Trim();

            if (rule.IndexOf("#content#") >= 0)
            {
                //rule = rule.Replace("#content#", "((?:(?!" + Regex.Split(rule, "#content#")[1] + ").)+)");        // 不包含结束标签的html
                rule = rule.Replace("#content#", "(.+?)");
            }
            if (rule.IndexOf("#url#") >= 0)
            {
                rule = rule.Replace("#url#", "([^<>\"']+)");
            }
            if (rule.IndexOf("#fileid#") >= 0)
            {
                rule = rule.Replace("#fileid#", "([^<>\"'\\.]+?)");
            }
            if (rule.IndexOf("#title#") >= 0)
            {
                rule = rule.Replace("#title#", "(.+?)");
            }
            if (rule.IndexOf("#keyword#") >= 0)
            {
                rule = rule.Replace("#keyword#", "(.+?)");
            }
            if (rule.IndexOf("#parameter#") >= 0)
            {
                rule = rule.Replace("#parameter#", "([^\\}]+)");
            }
            return rule;
        }


        /// <summary>
        /// 处理文章列表项
        /// 返回指定规则的文章编号 和标题
        /// </summary>
        /// <param name="strHTML">源代码</param>
        /// <param name="rule">规则</param>
        /// <returns></returns>
        public static string[] GetUrlAndTitle(string strHTML, string rule)
        {
           // rule = rule.Trim().ToLower().Replace("\n", "").Replace("\r", "").Replace("\t", "");
          
            if (rule.IndexOf("#FileID#") <= 0 || rule.IndexOf("#Title#") <= 0)
            {
                return null;
            }
            string[] tagArray = new string[2];
            tagArray[0] = GetHtmlContent(strHTML, rule.Replace("#Title#", ".+?"));
            tagArray[1] = GetHtmlContent(strHTML, rule.Replace("#FileID#", ".+?"));
            
            return tagArray;
        }




        /// <summary>
        /// 是否匹配
        /// </summary>
        /// <param name="strHTML"></param>
        /// <param name="rule"></param>
        /// <returns></returns>
        public static bool IsMatch(string strHTML, string rule)
        {
            if (strHTML == null || strHTML == "")
            {
                return false;
            }
            Regex r = new Regex(rule, RegexOptions.IgnoreCase);
            return r.IsMatch(strHTML);
        }

        /// <summary>
        /// 正则替换
        /// </summary>
        /// <param name="strHTML"></param>
        /// <param name="rule"></param>
        /// <returns></returns>
        public static string Replace(string orgStr, string rule, string str)
        {
            if (orgStr == null || orgStr == "")
            {
                return "";
            }
            Regex r = new Regex(rule, RegexOptions.IgnoreCase);
            return r.Replace(orgStr, str);
        }

         /// <summary>
        /// 过滤一些特殊html标签
        /// </summary>
        /// <param name="strHTML"></param>
        /// <returns></returns>
        public static string FilterHtml(string strHTML)
        {
            return Replace(strHTML, "<.+?>", "");
        }
        /// <summary>
        /// 过滤一些特殊html标签
        /// </summary>
        /// <param name="strHTML"></param>
        /// <returns></returns>
        public static string FilterHtmlLabel(string strHTML)
        {

            strHTML = strHTML.ToLower().Replace("<b>", "").Replace("</b>", "").Replace("<i>", "").Replace("</i>", "").Replace("<u>", "").Replace("</u>", "");      // 处理无属性的标签
            strHTML = strHTML.Replace("</a>", ""); // 先处理</a>

            int index = strHTML.IndexOf("<a");  // 查找a标签
            int endIndex = 0;
            while (index >= 0)
            {
                endIndex = strHTML.IndexOf(">", index);  // 查找结束标签

                if (endIndex > 0)
                {
                    strHTML = strHTML.Remove(index, endIndex - index + 1);
                }


                index = strHTML.IndexOf("<a");


            }


            index = strHTML.IndexOf("<img");  // 查找Img标签
            endIndex = 0;
            while (index >= 0)
            {
                endIndex = strHTML.IndexOf(">", index);  // 查找结束标签

                if (endIndex > 0)
                {
                    strHTML = strHTML.Remove(index, endIndex - index + 1);
                }


                index = strHTML.IndexOf("<img");


            }


            index = strHTML.IndexOf("<script");  // 查找script标签
            endIndex = 0;
            while (index >= 0)
            {
                endIndex = strHTML.IndexOf("</script>", index);  // 查找结束标签

                if (endIndex > 0)
                {
                    strHTML = strHTML.Remove(index, endIndex - index + 9);
                }


                index = strHTML.IndexOf("<script");


            }

            index = strHTML.IndexOf("<style");  // 查找script标签
            endIndex = 0;
            while (index >= 0)
            {
                endIndex = strHTML.IndexOf("</style>", index);  // 查找结束标签

                if (endIndex > 0)
                {
                    strHTML = strHTML.Remove(index, endIndex - index + 8);
                }


                index = strHTML.IndexOf("<style");


            }



            strHTML = strHTML.Replace("</font>", ""); // 先处理</font>
            index = strHTML.IndexOf("<font");  // 查找script标签
            endIndex = 0;
            while (index >= 0)
            {
                endIndex = strHTML.IndexOf(">", index);  // 查找结束标签

                if (endIndex > 0)
                {
                    strHTML = strHTML.Remove(index, endIndex - index + 1);
                }


                index = strHTML.IndexOf("<font");


            }
            /*** iframe 广告 **/
            strHTML = strHTML.Replace("</iframe>", ""); // 先处理</iframe>
            index = strHTML.IndexOf("<iframe");  // 查找iframe标签
            endIndex = 0;
            while (index >= 0)
            {
                endIndex = strHTML.IndexOf(">", index);  // 查找结束标签

                if (endIndex > 0)
                {
                    strHTML = strHTML.Remove(index, endIndex - index + 1);
                }
                index = strHTML.IndexOf("<iframe");
            }
            return strHTML;

        }


        /// <summary> 
        /// 将本地或者网络图片上传到指定的服务器(HttpWebRequest方法) 
        /// </summary> 
        /// <param name="address">文件上传到的服务器</param> 
        /// <param name="fileNamePath">要上传的本地文件（全路径）或者网络图片</param> 
        /// <param name="inputParamter">其他表单元素</param>
        /// <param name="cookie"></param>
        /// <returns>返回服务器的响应文本</returns>
        public static string UploadIamge(string url, string fileNamePath, Dictionary<string, string> inputDic, string setCookie, string fileKey, string refer, string encoding, out string location)
        {
            // 要上传的文件 
            byte[] bytes = null;
            location = "";
            string saveName = "";
            //fileNamePath = "C:/Users/Administrator/Desktop/1.jpg";
        
            // url += "?resp_charset=UTF8&pmlang=zh_CN&ef=if";
            if (url != null && url.Length > 0)
            {
                bytes = GetBinaryReader(fileNamePath);
                if (bytes == null)
                {
                    return "";
                }
                saveName = GetFileName(fileNamePath);
            }
            Encoding encode = Encoding.GetEncoding(encoding);

            //时间戳 
            string strBoundary = "----------"+DateTime.Now.Ticks.ToString("x");
            string end = "\r\n--" + strBoundary + "--\r\n";
          
            //请求头部信息 
            StringBuilder sb = new StringBuilder("");
            if (inputDic != null && inputDic.Count > 0)
            {
                //遍历字典取出表单普通空间的健和值
                foreach (KeyValuePair<string, string> dicItem in inputDic)
                {
                    sb.Append("--" + strBoundary);
                    sb.Append("\r\nContent-Disposition: form-data; name=\"" + dicItem.Key + "\"");
                    sb.Append("\r\n\r\n");
                    sb.Append(dicItem.Value);//value前面必须有2个换行
                    sb.Append("\r\n");
                }
            }
            sb.Append("--");
            sb.Append(strBoundary);
            if (bytes != null)
            {
                sb.AppendFormat("\r\nContent-Disposition: form-data; name=\"{0}\"; filename=\"", fileKey);
                sb.Append(saveName);
                sb.Append("\"\r\nContent-Type: application/octet-stream\r\n\r\n");
                
            }
            byte[] boundaryBytes = encode.GetBytes(end);
            string strPostHeader = sb.ToString();
            byte[] postHeaderBytes = encode.GetBytes(strPostHeader);

            // 根据uri创建HttpWebRequest对象 
            HttpWebRequest httpReq = null;
            //如果是发送HTTPS请求
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
            {
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;
                httpReq = WebRequest.Create(url) as HttpWebRequest;
            }
            else
            {
                httpReq = (HttpWebRequest)WebRequest.Create(url);
            }

            httpReq.Method = "POST";
            httpReq.KeepAlive = true;
            httpReq.Accept = "*/*";
            httpReq.UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.31 (KHTML, like Gecko) Chrome/26.0.1410.64 Safari/537.31"; //"Shockwave Flash";//
            
            /** httpReq.Headers.Add("Accept-Charset", "GBK,utf-8;q=0.7,*;q=0.3"); //:
             httpReq.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
             httpReq.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8");
            
            httpReq.AllowAutoRedirect = false;
            httpReq.Proxy = null;
            httpReq.ServicePoint.Expect100Continue = false;
            //是否使用 Nagle 不使用 提高效率 
            httpReq.ServicePoint.UseNagleAlgorithm = false;
            //最大连接数 
            httpReq.ServicePoint.ConnectionLimit = 65500;
            //数据是否缓冲 false 提高效率  
            httpReq.AllowWriteStreamBuffering = false;
            //对发送的数据不使用缓存 
            httpReq.AllowWriteStreamBuffering = false;*/
          // httpReq.Headers["Cookie"] =  setCookie == null ? "" : setCookie;
            //设置获得响应的超时时间（120秒） 
           // httpReq.Timeout = 120000;
            httpReq.ContentType = "multipart/form-data; boundary=" + strBoundary;
           // httpReq.Referer = refer == null ? "" : refer;
           // if (url.Contains("mp.weixin.qq.com/cgi-bin/filetransfer"))
           //  httpReq.Headers["Origin"] = "https://mp.weixin.qq.com";
            httpReq.ProtocolVersion = HttpVersion.Version11;
            long fileLength = bytes == null ? 0 : bytes.Length;
            long length = fileLength + postHeaderBytes.Length + boundaryBytes.Length;

            httpReq.ContentLength = length;

            try
            {

                Stream postStream = httpReq.GetRequestStream();

                //发送请求头部消息 
                postStream.Write(postHeaderBytes, 0, postHeaderBytes.Length);
                if (bytes != null && bytes.Length > 0)
                {
                    //System.Threading.Thread.Sleep(3000);
                    postStream.Write(bytes, 0, bytes.Length);
                    if(bytes.Length<200000){
                      //  System.Threading.Thread.Sleep(6000);
                    }
                }
                //添加尾部的时间戳 
                postStream.Write(boundaryBytes, 0, boundaryBytes.Length);
               

                //获取服务器端的响应 
                WebResponse webRespon = httpReq.GetResponse();
               // System.Threading.Thread.Sleep(3000);
                Stream s = webRespon.GetResponseStream();
                StreamReader sr = new StreamReader(s, encode);//Encoding.GetEncoding("utf-8")

                //读取服务器端返回的消息 
                String html = sr.ReadToEnd(); postStream.Close();
                if (webRespon.Headers != null && webRespon.Headers.Keys.Count > 0)
                {
                    try
                    {
                        string loc = webRespon.Headers["Location"];
                        if (loc != null && loc.Length > 0)
                        {
                            location = CreateUrl(url, loc);
                        }
                    }
                    catch (Exception) { location = ""; }
                }
                httpReq.Abort();
                sr.Close();
                webRespon.Close();
                return html;
            }
            catch (Exception e) { }
            finally
            {
                if (httpReq != null) httpReq.Abort();
            }
            return "";
        }
        private static  bool CheckValidationResult(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
          System.Security.Cryptography.X509Certificates.X509Chain chain, SslPolicyErrors errors)
        {
            return true; //总是接受
        } 

        public static string GetFileName(string path)
        {
            if (path == null || path.Length == 0)
            {
                return "";
            }
            int index = path.LastIndexOf("\\");
            if (index == -1)
            {
                index = path.LastIndexOf("/");
            }

            return path.Substring(index + 1);
        }
        /// <summary>
        /// 获取文件2进制流
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private static byte[] GetBinaryReader(string path)
        {
            try
            {
                byte[] bytes;
                if (IsMatch(path, "http://"))
                {
                    return new WebClient().DownloadData(path);
                   // return SendRequest_Img(path);
                }
                else
                {
                    FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                    bytes = new byte[(int)fs.Length];
                    fs.Read(bytes, 0, bytes.Length);
                    fs.Close();
                    return bytes;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static string Post(string url, string encoding, string postData,string setCookie,string refer )
        {
             HttpWebRequest myHttpWebRequest = null;
            
            try
            {
                //如果是发送HTTPS请求
                if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
                {
                    ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;
                    myHttpWebRequest = WebRequest.Create(url) as HttpWebRequest;
                }
                else
                {
                    myHttpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                }
                myHttpWebRequest.Method = "POST";
                myHttpWebRequest.KeepAlive = true ;
                myHttpWebRequest.ProtocolVersion = HttpVersion.Version10;
                myHttpWebRequest.Timeout = 60000;
                myHttpWebRequest.ContentType = "application/x-www-form-urlencoded";//application/x-www-form-urlencoded; charset=UTF-8
                myHttpWebRequest.Proxy = null;
                myHttpWebRequest.ServicePoint.Expect100Continue = false;
                //是否使用 Nagle 不使用 提高效率 
                myHttpWebRequest.ServicePoint.UseNagleAlgorithm = false;
                //最大连接数 
                myHttpWebRequest.ServicePoint.ConnectionLimit = 65500;
                //数据是否缓冲 false 提高效率  
                myHttpWebRequest.AllowWriteStreamBuffering = false;

                myHttpWebRequest.UserAgent = "Mozilla/5.0 (Windows NT 6.1; rv:23.0) Gecko/20100101 Firefox/23.0";
              //  myHttpWebRequest.Headers.Add("Accept-Language", "zh-cn,en-us;q=0.5");
                myHttpWebRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                if(refer.Contains("appmsgid=")){
                    myHttpWebRequest.Headers.Add("Cache-Control", "no-cache");
                    myHttpWebRequest.Headers.Add("Pragma", "no-cache");
                    myHttpWebRequest.Headers.Add("X-Requested-With", "XMLHttpRequest");		
                }
                 if(setCookie!=null)
                     myHttpWebRequest.Headers.Add("Cookie", setCookie);
            if (refer != null)
                myHttpWebRequest.Referer = refer;
                byte[] data = Encoding.GetEncoding(encoding).GetBytes(postData);
                myHttpWebRequest.ContentLength = data.Length;

                using (Stream reqStream = myHttpWebRequest.GetRequestStream())
                {
                    reqStream.Write(data, 0, data.Length);
                }

                // Get response
                HttpWebResponse response = (HttpWebResponse)myHttpWebRequest.GetResponse();
                Stream stream = response.GetResponseStream();
                stream.ReadTimeout = 10000;
                System.IO.StreamReader sr = new System.IO.StreamReader(stream, Encoding.GetEncoding(encoding));     //utf-8

               string strResult = sr.ReadToEnd();
                sr.Close();
                stream.Close();
                response.Close();
                return strResult;
               
            }
           catch (Exception e)
            {
                return null;
            }
            finally
            {
                if (myHttpWebRequest != null)
                {
                    myHttpWebRequest.Abort();
                }
            }
        }


        /// <summary>
        /// 根据原来url和loction 生成完整的url
        /// </summary>
        /// <param name="url"></param>
        /// <param name="loc"></param>
        /// <returns></returns>
        public static string CreateUrl(string url, string loc)
        {
            if (loc == null || loc.Length==0)
            {
                loc = "";
            }
            loc = loc.Replace("&amp;", "&");
            if (url == null || url.Length == 0)
            {
                return loc;
            }
            if (loc.ToLower().IndexOf("http") == 0) // 绝对路径
            {
                return loc;
            }
            else if (loc.ToLower().IndexOf("/") == 0) // 相对站点根目录路径
            {
                int index = url.IndexOf('/', 8);
                if (index >= 0)
                {
                    return url.Substring(0, index) + loc;
                }
                else
                {
                    return url + loc;
                }
            }
            else  // 相对路径
            {
                int index = url.Split('?')[0].LastIndexOf('/');
                if (index > 8)
                {
                    return url.Substring(0, (index + 1)) + loc;
                }
                else
                {
                    return url + "/" + loc;
                }
            }
        }

        /// <summary>
        /// 获取指定编码后的字符串
        /// </summary>
        /// <param name="str"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public static string GetEncoding(string str, string encoding)
        {
            if (encoding.ToLower() == "utf-8")
            {
                return UTF8(str);
            }
            return GB2312(str);
        }
        public static string GB2312(string str)
        {
            str = ReplaceSpecialChar(str);
            StringBuilder sb = new StringBuilder();
            Encoding en = Encoding.GetEncoding("GB2312");
            for (int i = 0; i < str.Length; i++)
            {
                byte[] byteCode = en.GetBytes(str[i].ToString());
                if (byteCode.Length == 2)
                {
                    sb.Append("%" + Convert.ToString(byteCode[0], 16) + "%" + Convert.ToString(byteCode[1], 16));
                }
                else
                {
                    sb.Append(str[i]);
                }
            }

            return sb.ToString();
        }


        /// <summary>
        /// 把汉字按utf-8 编码
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string UTF8(string str)
        {
            str = ReplaceSpecialChar(str);
            StringBuilder sb = new StringBuilder();
            Encoding en = Encoding.GetEncoding("UTF-8");
            for (int i = 0; i < str.Length; i++)
            {
                byte[] byteCode = en.GetBytes(str[i].ToString());
                if (byteCode.Length == 3)
                {
                    sb.Append("%" + Convert.ToString(byteCode[0], 16) + "%" + Convert.ToString(byteCode[1], 16) + "%" + Convert.ToString(byteCode[2], 16));
                }
                else
                {
                    sb.Append(str[i]);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        ///  替换特殊字符 '% &?='
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string ReplaceChar(string str)
        {
            return str.Replace("%", "%25").Replace(" ", "%20").Replace("&", "%26").Replace("?", "%3F").Replace("=", "%3D").Replace("#", "%23");
        }

        /// <summary>
        /// 替换特殊字符 '% &?:=/+@'
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string ReplaceSpecialChar(string str)
        {
            return str.Replace("%", "%25").Replace(" ", "%20").Replace("&", "%26").Replace("?", "%3F").Replace(":", "%3A").Replace("=", "%3D").Replace("/", "%2F").Replace("+", "%2B").Replace("@", "%40").Replace("#", "%23");
        }


        public static int ParseInt(string v)
        {
            try
            {
                return int.Parse(v);
            }
            catch (Exception){}
            return 0;
        }


        public static long ParseLong(string v)
        {
            try
            {
                return long.Parse(v);
            }
            catch (Exception) { }
            return 0;
        }
        /// <summary>
        /// 获取指定范围正随机数[1,10)
        /// </summary>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns></returns>
        public static int GetRandom(int min, int max)
        {
            if (min > max)
            {
                int temp = min;
                min = max;
                max = temp;
            }
            return new Random(GetRandomSeed()).Next(min, max);
        }
        /// <summary>
        /// 加密随机数生成器 生成随机种子
        /// </summary>
        /// <returns></returns>
        private static int GetRandomSeed()
        {
            byte[] bytes = new byte[4];
            System.Security.Cryptography.RNGCryptoServiceProvider rng = new System.Security.Cryptography.RNGCryptoServiceProvider();
            rng.GetBytes(bytes);
            return BitConverter.ToInt32(bytes, 0);
        }

    }
}
