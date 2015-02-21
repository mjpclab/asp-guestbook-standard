function TabControl(id) {
	var that = this;

	//Attributes of Objects
	var outerContainer = document.getElementById(id);
	var titleContainer;
	var pageContainer;
	var titles;
	var pages;

	//Attributes of Styles
	var outerContainerCssClass = '';
	var titleContainerCssClass = '';
	var pageContainerCssClass = '';
	var titleCssClass = '';
	var titleSelectedCssClass = '';
	var pageCssClass = '';

	//Other fields
	var selectedIndex = -1;
	that.savingFieldId = '';

	//Methods
	function init() {
		titleContainer = document.createElement('div');
		pageContainer = document.createElement('div');
		titles = [];
		pages = [];
	}

	function hookMethods() {
		//addPage
		that.addPage = function (id, caption) {
			var pageElement = document.getElementById(id);
			if (pageElement) {
				//Create caption
				var titleElement = document.createElement('div');
				titleElement.titleIndex = titles.length;
				titleElement.tabObject = that;
				titleElement.className = titleCssClass;
				titleElement.onclick = function () {
					var tab = this.tabObject;
					tab.selectPage(this.titleIndex);
					tab.savePageIndex();

					this.className += ( ' ' + titleSelectedCssClass);
				};

				//Add into titles[]
				titles[titles.length] = titleElement;
				titleElement.appendChild(document.createTextNode(caption));
				titleContainer.appendChild(titleElement);

				//Add page
				pageElement.className = pageCssClass;
				pageElement.style.display = 'none';
				pages[pages.length] = pageElement;
				pageContainer.appendChild(pageElement);

				//Show page
				if (selectedIndex == -1) this.selectPage(0);
				this.savePageIndex();
			}
		};

		// set???CssClass functions
		that.setTitleCssClass = function (className) {
			titleCssClass = className;
			for (var i = 0; i < titles.length; i++) {
				if (i === selectedIndex) {
					titles[i].className = titleCssClass + ' ' + titleSelectedCssClass;
				}
				else {
					titles[i].className = titleCssClass;
				}
			}
		};

		that.setTitleSelectedCssClass = function (className) {
			titleSelectedCssClass = className;
			if (selectedIndex > -1) {
				titles[selectedIndex].className = titleCssClass + ' ' + titleSelectedCssClass;
			}
		};

		that.setOuterContainerCssClass = function (className) {
			outerContainerCssClass = className;
			outerContainer.className = outerContainerCssClass;
		};

		that.setTitleContainerCssClass = function (className) {
			titleContainerCssClass = className;
			titleContainer.className = titleContainerCssClass;
		};

		that.setPageContainerCssClass = function (className) {
			pageContainerCssClass = className;
			pageContainer.className = pageContainerCssClass;
		};


		that.setPageCssClass = function (className) {
			pageCssClass = className;
			for (var i = 0; i < pages.length; i++) {
				pages[i].className = pageCssClass;
			}
		};

		//Others
		that.selectPage = function (newIndex) {
			if (newIndex > -1 && newIndex < pages.length && newIndex !== selectedIndex) {
				//Reset old page's CssClass & Style to titleStyle
				if (selectedIndex > -1) {
					titles[selectedIndex].className = titleCssClass;
					pages[selectedIndex].style.display = 'none';
				}

				//Overwrite new page'style by titleSelectedStyle
				selectedIndex = newIndex;
				titles[selectedIndex].className = titleCssClass + ' ' + titleSelectedCssClass;
				pages[selectedIndex].style.display = 'block';
			}
		};

		that.savePageIndex = function () {
			if (this.savingFieldId) {
				var savingField = document.getElementById(this.savingFieldId);
				if (savingField) savingField.value = selectedIndex;
			}
		};

		that.loadPageIndex = function () {
			if (this.savingFieldId) {
				var savingField = document.getElementById(this.savingFieldId);
				if (savingField) {
					var loadedIndex = parseInt(savingField.value);
					if (isFinite(loadedIndex)) this.selectPage(loadedIndex);
				}
			}
		};

		//Combine the Containers
		var fragment = document.createDocumentFragment();
		fragment.appendChild(titleContainer);
		fragment.appendChild(pageContainer);

		outerContainer.appendChild(fragment);
	}

	if (outerContainer) {
		init();
		hookMethods();
	}
	else {
		throw new Error('tab container not found')
	}
}