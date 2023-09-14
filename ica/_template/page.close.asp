<html>
<head></head>
<body>
<script>
if (top && top.opener && top.opener.location) {
	try {
		top.opener.location.reload();
	} finally {
	}
}
try {
	top.opener = 'Dummy';
	setTimeout("top.close()", 50);
} catch(err) {
	close(); 
}
</script>
</body>
</html>
<% Response.End %>
