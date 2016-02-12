function toArray(arrayObj) {
	var result;
	try {
		result = Array.prototype.slice.call(arrayObj);
	}
	catch (ex) {
		result = arrayObj;
	}

	return result;
}

function refresh_pagenum(increment) {
	function do_refresh_pagenum(increment) {
		var pageNums = document.getElementsByName('pagenum');
		pageNums = toArray(pageNums);
		if (pageNums.length) {
			var old_start = Number(pageNums[0].firstChild.data);
			var new_start = old_start + increment;

			if (new_start < 1) new_start = 1;
			if (new_start + PageListCount - 1 > PagesCount) new_start = PagesCount - PageListCount + 1;
			if (new_start < 1) new_start = 1;

			var delta = new_start - old_start;
			for (var i = 0; i < pageNums.length && delta != 0; i++) {
				var pageNum = pageNums[i];
				var old_current = Number(pageNum.firstChild.data);
				var new_current = old_current + delta;

				pageNum.href = pageNum.href.replace(/page=\d+/i, 'page=' + new_current.toString(10));
				pageNum.firstChild.data = new_current.toString(10);

				if (new_current === CurrentPage)
					pageNum.className = 'pagenum pagenum-current';
				else
					pageNum.className = 'pagenum';
			}
		}
	}

	do_refresh_pagenum(increment);
	ghandle = setInterval(function () {
		do_refresh_pagenum(increment);
	}, 250);
}

function stop_refresh_pagenum() {
	clearTimeout(ghandle);
}

//hide normal controls
var pageControls = document.getElementsByName('page-control');
pageControls = toArray(pageControls);
if (pageControls.length) {
	for (var i = 0; i < pageControls.length; i++) {
		pageControls[i].style.display = 'none';
	}
}

var jsPageControls = document.getElementsByName('js-page-control');
jsPageControls = toArray(jsPageControls);
if (jsPageControls.length) {
	for (var i = 0; i < jsPageControls.length; i++) {
		jsPageControls[i].style.display = 'block';
	}
}
