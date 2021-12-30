"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[2376],{3905:function(e,o,r){r.d(o,{Zo:function(){return u},kt:function(){return f}});var t=r(7294);function n(e,o,r){return o in e?Object.defineProperty(e,o,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[o]=r,e}function i(e,o){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var t=Object.getOwnPropertySymbols(e);o&&(t=t.filter((function(o){return Object.getOwnPropertyDescriptor(e,o).enumerable}))),r.push.apply(r,t)}return r}function a(e){for(var o=1;o<arguments.length;o++){var r=null!=arguments[o]?arguments[o]:{};o%2?i(Object(r),!0).forEach((function(o){n(e,o,r[o])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):i(Object(r)).forEach((function(o){Object.defineProperty(e,o,Object.getOwnPropertyDescriptor(r,o))}))}return e}function c(e,o){if(null==e)return{};var r,t,n=function(e,o){if(null==e)return{};var r,t,n={},i=Object.keys(e);for(t=0;t<i.length;t++)r=i[t],o.indexOf(r)>=0||(n[r]=e[r]);return n}(e,o);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(t=0;t<i.length;t++)r=i[t],o.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(n[r]=e[r])}return n}var l=t.createContext({}),k=function(e){var o=t.useContext(l),r=o;return e&&(r="function"==typeof e?e(o):a(a({},o),e)),r},u=function(e){var o=k(e.components);return t.createElement(l.Provider,{value:o},e.children)},s={inlineCode:"code",wrapper:function(e){var o=e.children;return t.createElement(t.Fragment,{},o)}},b=t.forwardRef((function(e,o){var r=e.components,n=e.mdxType,i=e.originalType,l=e.parentName,u=c(e,["components","mdxType","originalType","parentName"]),b=k(r),f=n,h=b["".concat(l,".").concat(f)]||b[f]||s[f]||i;return r?t.createElement(h,a(a({ref:o},u),{},{components:r})):t.createElement(h,a({ref:o},u))}));function f(e,o){var r=arguments,n=o&&o.mdxType;if("string"==typeof e||n){var i=r.length,a=new Array(i);a[0]=b;var c={};for(var l in o)hasOwnProperty.call(o,l)&&(c[l]=o[l]);c.originalType=e,c.mdxType="string"==typeof e?e:n,a[1]=c;for(var k=2;k<i;k++)a[k]=r[k];return t.createElement.apply(null,a)}return t.createElement.apply(null,r)}b.displayName="MDXCreateElement"},5158:function(e,o,r){r.r(o),r.d(o,{frontMatter:function(){return c},contentTitle:function(){return l},metadata:function(){return k},toc:function(){return u},default:function(){return b}});var t=r(7462),n=r(3366),i=(r(7294),r(3905)),a=["components"],c={},l=void 0,k={unversionedId:"The-Workbook-object/08-03 working-with-Workbook-references",id:"The-Workbook-object/08-03 working-with-Workbook-references",title:"08-03 working-with-Workbook-references",description:"working with Workbook references",source:"@site/docs/08-The-Workbook-object/08-03 working-with-Workbook-references.md",sourceDirName:"08-The-Workbook-object",slug:"/The-Workbook-object/08-03 working-with-Workbook-references",permalink:"/excel-vba-book/docs/The-Workbook-object/08-03 working-with-Workbook-references",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/08-The-Workbook-object/08-03 working-with-Workbook-references.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"08-02 understanding-the-Workbook-Object-Model",permalink:"/excel-vba-book/docs/The-Workbook-object/08-02 understanding-the-Workbook-Object-Model"},next:{title:"08-04 useful-properties-of-the-Workbook",permalink:"/excel-vba-book/docs/The-Workbook-object/08-04 useful-properties-of-the-Workbook"}},u=[{value:"working with Workbook references",id:"working-with-workbook-references",children:[],level:2}],s={toc:u};function b(e){var o=e.components,r=(0,n.Z)(e,a);return(0,i.kt)("wrapper",(0,t.Z)({},s,r,{components:o,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"working-with-workbook-references"},"working with Workbook references"),(0,i.kt)("p",null,"There are a couple of ways to obtain a reference to a Workbook that are useful:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"ActiveWorkbook - refers to the Workbook that has focus"),(0,i.kt)("li",{parentName:"ul"},"ThisWorkbook - refers to the Workbook which contains the code that is executing"),(0,i.kt)("li",{parentName:"ul"},"Workbooks.Open() - will open a Workbook and return a reference"),(0,i.kt)("li",{parentName:"ul"},"Workbooks(index) - will grab a reference to the currently opened Workbook"),(0,i.kt)("li",{parentName:"ul"},"Workbooks.Add() - will create a new blank Workbook or a Workbook according to a supplied template")),(0,i.kt)("p",null,"I find that all of those approaches are used equally across my code. The one exception might be ThisWorkbook which I typically avoid. In reality, I should probably use it more because I find myself going to some length to maintain a reference to a Workbook while opening or creating Workbooks."),(0,i.kt)("p",null,"For Workbooks, the biggest thing to be aware of that there are a number of unqualified references that exist within VBA that are a part of the ActiveWorkbook. Those include:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"Worksheets and Sheets"),(0,i.kt)("li",{parentName:"ul"},"Names?")),(0,i.kt)("p",null,"These unqualified references can really bite you when you are expecting it. The problem with unqualified references is that they work great initially, before the workflow becomes complex. They will then silently fail later when you start creating new Workbooks and otherwise changing the focus or active Workbook. The problem is that nearly all of the unqualified references apply to the ActiveWorkbook. Working with Workbooks is the one task that will often change the focus of Excel regardless of how you create things."))}b.isMDXComponent=!0}}]);