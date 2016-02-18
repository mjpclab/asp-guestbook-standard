(function enableRefreshCaptcha() {
	var captcha = document.getElementById('captcha');
	if(captcha) {
		captcha.onclick = function () {
			this.src = this.src.replace(/t=\d*/, 't=' + (new Date()).getTime());
		};
	}
}());
