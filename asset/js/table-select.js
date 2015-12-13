(function enableCheckAll() {
	$('table thead input[type=checkbox]').change(function () {
		var name = this.name;
		var checked = this.checked;

		var $checkBoxes = $('input[name="' + name + '"][type=checkbox]', this.form).not(this);
		$checkBoxes.prop('checked', checked);
	});
}());
