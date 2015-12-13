(function enableRefreshCaptcha() {
	var captcha = document.getElementById('captcha');
	captcha.onclick = function () {
		this.src = this.src.replace(/t=\d*/, 't=' + (new Date()).getTime());
	};
}());
