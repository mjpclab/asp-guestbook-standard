(function () {
	if (!document.querySelector) return;

	function checkAllHandler() {
		var name = this.name;
		var checked = this.checked;

		var checkboxes = document.querySelectorAll('input[name="' + name + '"][type=checkbox]');
		for (var i = 0; i < checkboxes.length; i++) {
			checkboxes[i].checked = checked;
		}
	}

	var headerChecks = document.querySelectorAll('table thead input[type=checkbox]');
	for (var i = 0; i < headerChecks.length; i++) {
		var headerCheck = headerChecks[i];
		headerCheck.onchange = checkAllHandler;
	}
}());