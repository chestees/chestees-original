<script src="/js/functions.js" type="text/javascript"></script>
<meta name="robots" content="index, follow" />
<meta name="description" content="<%=cDescription%>" />
<meta name="keywords" content="<%If strKeywords <> "" Then Response.Write(strKeywords & " ")%><%=cKeywords%>" />
<link rel="stylesheet" href="/css/style.css?v=1">
<!--[if IE 6]>
	<link rel="stylesheet" type="text/css" href="/css/ie6.css" />
<![endif]-->
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>
<body>
<script type="text/javascript">
	//GOOGLE PLUS
	(function () {
		var po = document.createElement('script'); po.type = 'text/javascript'; po.async = true;
		po.src = 'https://apis.google.com/js/plusone.js';
		var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(po, s);
	})();
	
	var _gaq = _gaq || [];
	_gaq.push(['_setAccount', 'UA-3725315-1']);
	_gaq.push(['_setDomainName', '.chestees.com']);
	_gaq.push(['_trackPageview']);
	
	(function() {
	var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
	ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
	var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
	})();
</script>
<!-- Google Analytics Social Button Tracking -->
<script src="/js/ga_social_tracking.js"></script>
<!-- Load Twitter JS-API asynchronously -->
<script type="text/javascript" charset="utf-8">
	window.twttr = (function (d, s, id) {
		var t, js, fjs = d.getElementsByTagName(s)[0];
		if (d.getElementById(id)) return; js = d.createElement(s); js.id = id;
		js.src = "//platform.twitter.com/widgets.js"; fjs.parentNode.insertBefore(js, fjs);
		return window.twttr || (t = { _e: [], ready: function (f) { t._e.push(f) } });
	} (document, "script", "twitter-wjs"));

	// Wait for the asynchronous resources to load
	twttr.ready(function (twttr) {
		_ga.trackTwitter(); //Google Analytics tracking
	});
</script>
<div id="fb-root"></div>
<script>(function (d, s, id) {
    var js, fjs = d.getElementsByTagName(s)[0];
    if (d.getElementById(id)) { return; }
    js = d.createElement(s); js.id = id;
    js.src = "//connect.facebook.net/en_US/all.js#xfbml=1&appId=230800510267354";
    fjs.parentNode.insertBefore(js, fjs);
  } (document, 'script', 'facebook-jssdk'));</script>
<div id="MessageBar"></div>
<% Call Header() %>