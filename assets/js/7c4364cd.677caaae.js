"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[9790],{3905:function(e,t,r){r.d(t,{Zo:function(){return u},kt:function(){return f}});var o=r(7294);function n(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function i(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,o)}return r}function s(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?i(Object(r),!0).forEach((function(t){n(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):i(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function a(e,t){if(null==e)return{};var r,o,n=function(e,t){if(null==e)return{};var r,o,n={},i=Object.keys(e);for(o=0;o<i.length;o++)r=i[o],t.indexOf(r)>=0||(n[r]=e[r]);return n}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(o=0;o<i.length;o++)r=i[o],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(n[r]=e[r])}return n}var c=o.createContext({}),l=function(e){var t=o.useContext(c),r=t;return e&&(r="function"==typeof e?e(t):s(s({},t),e)),r},u=function(e){var t=l(e.components);return o.createElement(c.Provider,{value:t},e.children)},h={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},p=o.forwardRef((function(e,t){var r=e.components,n=e.mdxType,i=e.originalType,c=e.parentName,u=a(e,["components","mdxType","originalType","parentName"]),p=l(r),f=n,k=p["".concat(c,".").concat(f)]||p[f]||h[f]||i;return r?o.createElement(k,s(s({ref:t},u),{},{components:r})):o.createElement(k,s({ref:t},u))}));function f(e,t){var r=arguments,n=t&&t.mdxType;if("string"==typeof e||n){var i=r.length,s=new Array(i);s[0]=p;var a={};for(var c in t)hasOwnProperty.call(t,c)&&(a[c]=t[c]);a.originalType=e,a.mdxType="string"==typeof e?e:n,s[1]=a;for(var l=2;l<i;l++)s[l]=r[l];return o.createElement.apply(null,s)}return o.createElement.apply(null,r)}p.displayName="MDXCreateElement"},7865:function(e,t,r){r.r(t),r.d(t,{frontMatter:function(){return a},contentTitle:function(){return c},metadata:function(){return l},toc:function(){return u},default:function(){return p}});var o=r(7462),n=r(3366),i=(r(7294),r(3905)),s=["components"],a={},c=void 0,l={unversionedId:"The-Workbook-object/08-04 useful-properties-of-the-Workbook",id:"The-Workbook-object/08-04 useful-properties-of-the-Workbook",title:"08-04 useful-properties-of-the-Workbook",description:"useful properties of the Workbook",source:"@site/docs/08-The-Workbook-object/08-04 useful-properties-of-the-Workbook.md",sourceDirName:"08-The-Workbook-object",slug:"/The-Workbook-object/08-04 useful-properties-of-the-Workbook",permalink:"/docs/The-Workbook-object/08-04 useful-properties-of-the-Workbook",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/08-The-Workbook-object/08-04 useful-properties-of-the-Workbook.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"08-03 working-with-Workbook-references",permalink:"/docs/The-Workbook-object/08-03 working-with-Workbook-references"},next:{title:"The Application object",permalink:"/docs/The-Application-object/09 The-Application-object"}},u=[{value:"useful properties of the Workbook",id:"useful-properties-of-the-workbook",children:[{value:"Worksheets vs. Sheets",id:"worksheets-vs-sheets",children:[],level:3}],level:2}],h={toc:u};function p(e){var t=e.components,r=(0,n.Z)(e,s);return(0,i.kt)("wrapper",(0,o.Z)({},h,r,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"useful-properties-of-the-workbook"},"useful properties of the Workbook"),(0,i.kt)("p",null,"Although I have railed against the Workbook object, there are a handful of things that it can do:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"Reference ",(0,i.kt)("inlineCode",{parentName:"li"},"Names")," which contains all of the global named ranges"),(0,i.kt)("li",{parentName:"ul"},"Others?"),(0,i.kt)("li",{parentName:"ul"},"Charts?")),(0,i.kt)("h3",{id:"worksheets-vs-sheets"},"Worksheets vs. Sheets"),(0,i.kt)("p",null,"WHen working with Worksheets, there are a pair of objects which will provide access to the underlying Sheets. They are different in how they handle Charts which are visible as a Worksheet. The rule is: Sheets will return the Charts, whereas Worksheets will only return the list of objects which are actually Worksheets. If you do not use Charts as Worksheets, then you will never notice a difference between these two objects. The one thing you will notice is that the ActiveWorksheet will not be of type Worksheet which means that you can never get Intellisense on one of the most useful objects."))}p.isMDXComponent=!0}}]);