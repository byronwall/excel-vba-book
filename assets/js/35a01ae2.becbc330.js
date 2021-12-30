"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[5270],{3905:function(e,n,t){t.d(n,{Zo:function(){return s},kt:function(){return h}});var a=t(7294);function o(e,n,t){return n in e?Object.defineProperty(e,n,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[n]=t,e}function i(e,n){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);n&&(a=a.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),t.push.apply(t,a)}return t}function r(e){for(var n=1;n<arguments.length;n++){var t=null!=arguments[n]?arguments[n]:{};n%2?i(Object(t),!0).forEach((function(n){o(e,n,t[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):i(Object(t)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(t,n))}))}return e}function d(e,n){if(null==e)return{};var t,a,o=function(e,n){if(null==e)return{};var t,a,o={},i=Object.keys(e);for(a=0;a<i.length;a++)t=i[a],n.indexOf(t)>=0||(o[t]=e[t]);return o}(e,n);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(a=0;a<i.length;a++)t=i[a],n.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(o[t]=e[t])}return o}var c=a.createContext({}),l=function(e){var n=a.useContext(c),t=n;return e&&(t="function"==typeof e?e(n):r(r({},n),e)),t},s=function(e){var n=l(e.components);return a.createElement(c.Provider,{value:n},e.children)},u={inlineCode:"code",wrapper:function(e){var n=e.children;return a.createElement(a.Fragment,{},n)}},p=a.forwardRef((function(e,n){var t=e.components,o=e.mdxType,i=e.originalType,c=e.parentName,s=d(e,["components","mdxType","originalType","parentName"]),p=l(t),h=o,f=p["".concat(c,".").concat(h)]||p[h]||u[h]||i;return t?a.createElement(f,r(r({ref:n},s),{},{components:t})):a.createElement(f,r({ref:n},s))}));function h(e,n){var t=arguments,o=n&&n.mdxType;if("string"==typeof e||o){var i=t.length,r=new Array(i);r[0]=p;var d={};for(var c in n)hasOwnProperty.call(n,c)&&(d[c]=n[c]);d.originalType=e,d.mdxType="string"==typeof e?e:o,r[1]=d;for(var l=2;l<i;l++)r[l]=t[l];return a.createElement.apply(null,r)}return a.createElement.apply(null,t)}p.displayName="MDXCreateElement"},2926:function(e,n,t){t.r(n),t.d(n,{frontMatter:function(){return d},contentTitle:function(){return c},metadata:function(){return l},toc:function(){return s},default:function(){return p}});var a=t(7462),o=t(3366),i=(t(7294),t(3905)),r=["components"],d={},c=void 0,l={unversionedId:"overview-of-building-an-addin/14-01 introduction-to-creating-an-addin",id:"overview-of-building-an-addin/14-01 introduction-to-creating-an-addin",title:"14-01 introduction-to-creating-an-addin",description:"introduction to creating an addin",source:"@site/docs/14-overview-of-building-an-addin/14-01 introduction-to-creating-an-addin.md",sourceDirName:"14-overview-of-building-an-addin",slug:"/overview-of-building-an-addin/14-01 introduction-to-creating-an-addin",permalink:"/docs/overview-of-building-an-addin/14-01 introduction-to-creating-an-addin",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/14-overview-of-building-an-addin/14-01 introduction-to-creating-an-addin.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"overview of building an addin",permalink:"/docs/overview-of-building-an-addin/14 overview-of-building-an-addin"},next:{title:"14-02 creating-an-addin",permalink:"/docs/overview-of-building-an-addin/14-02 creating-an-addin"}},s=[{value:"introduction to creating an addin",id:"introduction-to-creating-an-addin",children:[],level:2}],u={toc:s};function p(e){var n=e.components,t=(0,o.Z)(e,r);return(0,i.kt)("wrapper",(0,a.Z)({},u,t,{components:n,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"introduction-to-creating-an-addin"},"introduction to creating an addin"),(0,i.kt)("p",null,"This chapter will focus on creating an addin for Excel using VBA. There are other ways to create an addin but using VBA is simple because it can be done entirely from Excel and the Visual Basic Editor. The main distinction between an addin and other VBA code is that an addin is meant to be available to all open Workbooks without having to put the code inside a Workbook. This can be a very nice thing to have if you are regularly do the same or similar operations across different Workbooks. The alternative to an addin is often to maintain a library of code that you regularly export/import into macro enabled files as needed. This can create a mess as you change code in one file but not in another. The alternative also typically requires you to put the code inside a the Workbook and make it macro enabled. For certain applications, this is a non-starter. The one other alternative to a true addin is to create a Workbook that contains the code you want, and then you can open that file and execute the code in the context of whatever other files are open. This works, and creating an addin can be viewed as the logical conclusion of this approach. More than the logical conclusion, this is actually the first step for creating an addin."),(0,i.kt)("p",null,"When considering whether or not to create a proper addin with your code, consider the following:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"An addin provides a nice package for helper code and UDFs that might be used in multiple places"),(0,i.kt)("li",{parentName:"ul"},"An addin has easy access to the Ribbon and can create its own Ribbon tab"),(0,i.kt)("li",{parentName:"ul"},"An addin can be put in a central location and used as a repository of code for an organization (works best if the file is read-only)")),(0,i.kt)("p",null,"Item 1 in the list above is typically enough of a reason to consider creating an addin. A common example of an addin is as a personal repository of VBA code. This typically replaces the use of the Personal Workbook, which I have never found to work well."),(0,i.kt)("p",null,"When considering a personal addin, one of the biggest upsides is that you can always open the VBE and have immediate access to your library of code. This makes it easy to make edits and save the new addin. Immediately, your updated code is available for future use in all your Workbooks."),(0,i.kt)("p",null,"There are a couple of downsides related to addins:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"UDFs from an addin require that anyone opening the spreadsheet has the addin loaded"),(0,i.kt)("li",{parentName:"ul"},"For code in a single Workbook, it is often easier to simply use a macro enabled Workbook and save the code directly there"),(0,i.kt)("li",{parentName:"ul"},'Some folks are highly resistant to "installing an addin" but will happily open a XLSM file. These are equivalent in the case of opening an addin, but the hesitation still exists.')),(0,i.kt)("p",null,"Point 2 above is worth expanding on. Sometimes it's tempting to add code to an existing addin that make sense only in the context of a single file. This works well if you and everyone else have the addin. This starts to become a nuisance when you are constantly going through your addin to find code that should have been place in a Workbook to start. The cleaner way to store code that may be useful later is to place a copy of it in a personal addin. This ensures that the original code is always available in the Workbook and that future updates to the code don't break the original application."))}p.isMDXComponent=!0}}]);