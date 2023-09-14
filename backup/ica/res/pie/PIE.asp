<%

'This file is a wrapper, for use in PHP environments, which serves PIE.htc using the
'correct content-type, so that IE will recognize it as a behavior.  Simply specify the
'behavior property to fetch this .php file instead of the .htc directly:

'.myElement {
'    [ ...css3 properties... ]
'    behavior: url(PIE.php);
'}

'This is only necessary when the web server is not configured to serve .htc files with
'the text/x-component content-type, and cannot easily be configured to do so (as is the
'case with some shared hosting providers).

' header( 'Content-type: text/x-component' );
Response.AddHeader "Content-type","text/x-component"
%>
<PUBLIC:COMPONENT lightWeight="true">
<!-- saved from url=(0014)about:internet -->
<PUBLIC:ATTACH EVENT="oncontentready" FOR="element" ONEVENT="init()" />
<PUBLIC:ATTACH EVENT="ondocumentready" FOR="element" ONEVENT="init()" />
<PUBLIC:ATTACH EVENT="ondetach" FOR="element" ONEVENT="cleanup()" />
<script type="text/javascript">
	var d = element, g = d.document, j = g.documentMode || 0;
	!window.PIE && j < 10 && function () {
		var a = {}, k, i, b, l, h;
		window.PIE = {
			attach: function (c) { a[c.uniqueID] = c },
			detach: function (c) { delete a[c.uniqueID] }
		};
		b = g.createElement("div");
		b.innerHTML = "<!--[if IE 6]><i></i><![endif]-->";
		l = b.getElementsByTagName("i")[0]; if (b = g.location.href.match(/pie-load-path=([^&]+)/)) b = decodeURIComponent(b[1]);
		b || (b = g.documentElement.currentStyle.getAttribute((l ? "" : "-") + "pie-load-path"));
		if (!b) {
			k = /BEHAVIOR: url\(([^\)]*PIE[^\)]*)/i;
			i = function (c) {
				for (var f = c.length, e; f--; )
					if (e = (e = c[f].cssText.match(k)) ? e[1].substring(0, e[1].lastIndexOf("/")) : i(c[f].imports))
						break;
				return e
			};
			b = i(g.styleSheets)
		}
		if (b) {
			h = g.createElement("script");
			h.onreadystatechange = function () {
				var c = window.PIE, f = h.readyState, e;
				if (a && (f === "complete" || f === "loaded"))
					if ("version" in c) {
						for (e in a)
							a.hasOwnProperty(e) && c.attach(a[e]);
						a = 0
					}
			};
			h.src = "/res/pie/PIE_IE" + (j < 9 ? "678" : "9") + ".js";

			(g.getElementsByTagName("head")[0] || g.body).appendChild(h)
		}
	} ();  
	function init() { if (g.media !== "print") { var a = window.PIE; a && a.attach(d) } }
	function cleanup() { if (g.media !== "print") { var a = window.PIE; a && a.detach(d) } d = 0 } d.readyState === "complete" && init();

</script>

<script type="text/vbscript"></script>
</PUBLIC:COMPONENT>