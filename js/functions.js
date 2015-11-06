$(document).ready(function() {
	$('#sizing').click(function() {
		popUpWindow('/sizing-info/','40','20','520','425');
	});
	
	$('#socialFB').click(function() {
		popUpWindow($(this).attr('blerg'),'40','20','620','425');
	});
});
var popUpWin=0;
function popUpWindow(URLStr, left, top, width, height)
{
	if(popUpWin) {
		if(!popUpWin.closed) popUpWin.close();
	}
	popUpWin = open(URLStr, 'popUpWin', 'toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=1,resizable=yes,copyhistory=yes,width='+width+',height='+height+',left='+left+', top='+top+',screenX='+left+',screenY='+top+'');
}