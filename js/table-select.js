(function () {
	function checkAllHandler() {
		var name = this.name;
		var checked = this.checked;

		var $checkBoxes = $('input[name="' + name + '"][type=checkbox]',this.form).not(this);
		$checkBoxes.each(function (index, checkBox) {
			checkBox.checked = checked;
		});
	}

	var $headerChecks = $('table thead input[type=checkbox]');
	$headerChecks.change(checkAllHandler);
}());
