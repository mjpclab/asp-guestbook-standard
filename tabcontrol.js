
function TabControl(id)
{
	// A Page Control Javascript code
	// powered by MJ PC Lab, all rights reserved.
	// http://mjpclab.net/
	// Version 1.31
	// Update: Oct 2006

	//Attributes of Objects
	this.outerContainer=document.getElementById(id);
	this.titleContainer=document.createElement('div');
	this.pageContainer=document.createElement('div');
	this.titles=new Array();
	this.pages=new Array();
	
	//Attributes of Styles
	this.outerContainerCssClass='';
	this.titleContainerCssClass='';
	this.pageContainerCssClass='';
	this.titleCssClass='';
	this.titleSelectedCssClass='';
	this.titleOverCssClass='';
	this.titleOverSelectedCssClass='';
	this.pageCssClass='';

	this.outerContainerStyle=new Array();
	this.titleContainerStyle=new Array();
	this.pageContainerStyle=new Array();
	this.titleStyle=new Array();
	this.titleSelectedStyle=new Array();
	this.titleOverStyle=new Array();
	this.titleOverSelectedStyle=new Array();
	this.pageStyle=new Array();
	
	//Other fields
	this.selectedIndex=-1;
	this.savingFieldId='';
	
	//Methods
	this.hookMethods=function()
	{
		if(this.outerContainer)
		{
			//addPage
			this.addPage=function(id,caption)
			{
				var e=document.getElementById(id);
				if(e)
				{
					//Create caption
					var divTitle=document.createElement('div');
					for(var i in this.titleStyle) divTitle.style[i]=this.titleStyle[i];
					divTitle.titleIndex=this.titles.length;
					divTitle.parentClass=this;
					divTitle.onclick=function()
					{
						//Here, "this" means the divTitle
						var tab=this.parentClass
						tab.selectPage(this.titleIndex);
						tab.savePageIndex();
						
						this.className=tab.getTitleOverSelectedCssClass();
						for(var i in tab.titleOverSelectedStyle) this.style[i]=tab.titleOverSelectedStyle[i];
					}
					divTitle.onmouseover=function()
					{
						//Here, "this" means the divTitle
						this.className=this.parentClass.getTitleOverCssClass();
						for(var i in this.parentClass.titleOverStyle) this.style[i]=this.parentClass.titleOverStyle[i];
						if(this.titleIndex==this.parentClass.selectedIndex)
						{
							this.className=this.parentClass.getTitleOverSelectedCssClass();
							for(var i in this.parentClass.titleOverSelectedStyle) this.style[i]=this.parentClass.titleOverSelectedStyle[i];
						}
					}
					divTitle.onmouseout=function()
					{
						//Here, "this" means the divTitle
						this.className=this.parentClass.getTitleCssClass();
						this.parentClass.clearTitleStyle(this);
						for(var i in this.parentClass.titleStyle) this.style[i]=this.parentClass.titleStyle[i];
						if(this.titleIndex==this.parentClass.selectedIndex)
						{
							this.className=this.parentClass.getTitleSelectedCssClass();
							for(var i in this.parentClass.titleSelectedStyle) this.style[i]=this.parentClass.titleSelectedStyle[i];
						}
					}
					
					//Add into titles[]
					this.titles[this.titles.length]=divTitle;
					divTitle.appendChild(document.createTextNode(caption));
					this.titleContainer.appendChild(divTitle);
					
					//Add page
					for(var i in this.pageStyle) e.style[i]=this.pageStyle[i];
					this.pages[this.pages.length]=e;
					this.pageContainer.appendChild(e);
					
					//Show page
					if(this.selectedIndex==-1) this.selectPage(0);
					this.loadPageIndex();
				}
			}
			
			
			//getTitle???CssClass
			this.getTitleCssClass=				function() {return this.titleCssClass;};
			this.getTitleOverCssClass=			function() {return this.titleCssClass + ' ' + this.titleOverCssClass;};
			this.getTitleOverSelectedCssClass=	function() {return this.titleCssClass + ' ' + this.titleSelectedCssClass + ' ' + this.titleOverCssClass + ' ' + this.titleOverSelectedCssClass;};
			this.getTitleSelectedCssClass=		function() {return this.titleCssClass + ' ' + this.titleSelectedCssClass;};
			
			// set???CssClass functions
			this.setOuterContainerCssClass=function(className)
			{
				this.outerContainerCssClass=className;
				this.outerContainer.className=className;
			}

			this.setTitleContainerCssClass=function(className)
			{
				this.titleContainerCssClass=className;
				this.titleContainer.className=className;
			}

			this.setPageContainerCssClass=function(className)
			{
				this.pageContainerCssClass=className;
				this.pageContainer.className=className;
			}
			
			this.setTitleCssClass=function(className)
			{
				this.titleCssClass=className;
				for(var i=0;i<this.titles.length;i++)
				{
					this.titles[i].className=this.getTitleCssClass();
					if(i==this.selectedIndex)this.titles[i].className=this.getTitleSelectedCssClass();
				}
			}
			
			this.setTitleOverCssClass=function(className)
			{
				this.titleOverCssClass=className;
			}
			
			this.setTitleOverSelectedCssClass=function(className)
			{
				this.titleOverSelectedCssClass=className;
			}

			this.setTitleSelectedCssClass=function(className)
			{
				this.titleSelectedCssClass=className;
				if(this.selectedIndex>-1)this.titles[this.selectedIndex].className=this.getTitleSelectedCssClass();
			}

			this.setPageCssClass=function(className)
			{
				this.pageCssClass=className;
				for(var i=0;i<this.pages.length;i++)
				{
					if(this.pages[i]) this.pages[i].className=className;
				}
			}

			//set???Style functions
			this.setOuterContainerStyle=function(attr,value)
			{
				this.outerContainerStyle[attr]=value;
				this.outerContainer.style[attr]=value;
			}

			this.setTitleContainerStyle=function(attr,value)
			{
				this.titleContainerStyle[attr]=value;
				this.titleContainer.style[attr]=value;
			}

			this.setPageContainerStyle=function(attr,value)
			{
				this.pageContainerStyle[attr]=value;
				this.pageContainer.style[attr]=value;
			}
			
			this.setTitleStyle=function(attr,value)
			{
				this.titleStyle[attr]=value;
				for(var i=0;i<this.titles.length;i++)
				{
					if(this.titles[i] && this.titles[i].style)
					{
						this.titles[i].style[attr]=value;
						if(i==this.selectedIndex && this.titleSelectedStyle[attr])this.titles[i].style[attr]=this.titleSelectedStyle[attr];
					}
				}
			}
			
			this.setTitleOverStyle=function(attr,value)
			{
				this.titleOverStyle[attr]=value;
			}
			
			this.setTitleOverSelectedStyle=function(attr,value)
			{
				this.titleOverSelectedStyle[attr]=value;
			}

			this.setTitleSelectedStyle=function(attr,value)
			{
				this.titleSelectedStyle[attr]=value;
				if(this.selectedIndex>-1)this.titles[this.selectedIndex].style[attr]=value;
			}

			this.setPageStyle=function(attr,value)
			{
				this.pageStyle[attr]=value;
				for(var i=0;i<this.pages.length;i++)
				{
					if(this.pages[i] && this.pages[i].style) this.pages[i].style[attr]=value;
				}
			}
			
			this.clearTitleStyle=function(divTitle)
			{
				if(divTitle && divTitle.style)
				{
					for(var i in this.titleStyle) divTitle.style[i]='';
					for(var i in this.titleSelectedStyle) divTitle.style[i]='';
					for(var i in this.titleOverStyle) divTitle.style[i]='';
					for(var i in this.titleOverSelectedStyle) divTitle.style[i]='';
				}
			}
			
			
			//Others
			this.selectPage=function(index)
			{
				if(index>-1 && index<this.pages.length && index!=this.selectedIndex)
				{
					//Reset old page's CssClass & Style to titleStyle
					if(this.selectedIndex>-1)
					{
						this.titles[this.selectedIndex].className=this.titleCssClass;
						
						this.clearTitleStyle(this.titles[this.selectedIndex]);
						for(var i in this.titleStyle) this.titles[this.selectedIndex].style[i]=this.titleStyle[i];
						this.pages[this.selectedIndex].style['display']='none';
					}
					
					//Overwrite new page'style by titleSelectedStyle
					this.selectedIndex=index;
					this.titles[this.selectedIndex].className=this.getTitleSelectedCssClass();
					for(var i in this.titleSelectedStyle) this.titles[this.selectedIndex].style[i]=this.titleSelectedStyle[i];
					this.pages[this.selectedIndex].style['display']='block';
				}
			}
			
			this.savePageIndex=function()
			{
				var savingField=document.getElementById(this.savingFieldId);
				if(savingField && typeof(savingField.value).toLowerCase()=='string') savingField.value=this.selectedIndex;
			}
			
			this.loadPageIndex=function()
			{
				var savingField=document.getElementById(this.savingFieldId);
				if(savingField && savingField.value)
				{
					var loadedIndex=parseInt(savingField.value);
					if(!isNaN(loadedIndex)) this.selectPage(loadedIndex);
				}
			}
			
			this.preloadImage=function(url)
			{
				var img=document.createElement('img');
				img.alt='';
				img.src=url;
			}
			
			
			
			//Initialize default values			
			//this.setTitleContainerStyle('width','100%');
			this.setTitleContainerStyle('styleFloat','left');
			this.setTitleContainerStyle('cssFloat','left');
			this.setTitleContainerStyle('clear','both');
		
			this.setPageContainerStyle('width','100%');
			this.setPageContainerStyle('styleFloat','left');
			this.setPageContainerStyle('cssFloat','left');
			this.setPageContainerStyle('clear','both');
			
			this.setTitleStyle('styleFloat','left');
			this.setTitleStyle('cssFloat','left');
			this.setTitleStyle('clear','none');
			this.setTitleStyle('cursor','pointer');
			
			this.setPageStyle('width','100%');
			this.setPageStyle('display','none');

			//Combine the Containers
			this.outerContainer.appendChild(this.titleContainer);
			this.outerContainer.appendChild(this.pageContainer);
		}
	}
	
	this.about=function()
	{
		alert(
			'A JavaScript Non-Refresh Tab-Page Control Class\n' +
			'Powered by MJ PC Lab\n' +
			'http://mjpclab.net/\n' +
			'2006.10\n' +
			''
		);
	}
	
	if(this.outerContainer) this.hookMethods();
}