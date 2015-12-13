(function enableCheckAll() {
	$('table thead input.checkbox').change(function () {
		var name = this.name;
		var checked = this.checked;

		var $checkBoxes = $('input.' + name, this.form).not(this);
		$checkBoxes.prop('checked', checked);
	});
}());
