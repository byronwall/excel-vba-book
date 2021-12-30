"use strict";(self.webpackChunksite=self.webpackChunksite||[]).push([[8645],{3905:function(e,t,r){r.d(t,{Zo:function(){return s},kt:function(){return m}});var n=r(7294);function a(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function o(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,n)}return r}function l(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?o(Object(r),!0).forEach((function(t){a(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):o(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function i(e,t){if(null==e)return{};var r,n,a=function(e,t){if(null==e)return{};var r,n,a={},o=Object.keys(e);for(n=0;n<o.length;n++)r=o[n],t.indexOf(r)>=0||(a[r]=e[r]);return a}(e,t);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(n=0;n<o.length;n++)r=o[n],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(a[r]=e[r])}return a}var c=n.createContext({}),u=function(e){var t=n.useContext(c),r=t;return e&&(r="function"==typeof e?e(t):l(l({},t),e)),r},s=function(e){var t=u(e.components);return n.createElement(c.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},h=n.forwardRef((function(e,t){var r=e.components,a=e.mdxType,o=e.originalType,c=e.parentName,s=i(e,["components","mdxType","originalType","parentName"]),h=u(r),m=a,f=h["".concat(c,".").concat(m)]||h[m]||p[m]||o;return r?n.createElement(f,l(l({ref:t},s),{},{components:r})):n.createElement(f,l({ref:t},s))}));function m(e,t){var r=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var o=r.length,l=new Array(o);l[0]=h;var i={};for(var c in t)hasOwnProperty.call(t,c)&&(i[c]=t[c]);i.originalType=e,i.mdxType="string"==typeof e?e:a,l[1]=i;for(var u=2;u<o;u++)l[u]=r[u];return n.createElement.apply(null,l)}return n.createElement.apply(null,r)}h.displayName="MDXCreateElement"},3515:function(e,t,r){r.r(t),r.d(t,{frontMatter:function(){return i},contentTitle:function(){return c},metadata:function(){return u},toc:function(){return s},default:function(){return h}});var n=r(7462),a=r(3366),o=(r(7294),r(3905)),l=["components"],i={},c=void 0,u={unversionedId:"cheat-sheets/01-03 Excel-Object-Model-Cheat-Sheet",id:"cheat-sheets/01-03 Excel-Object-Model-Cheat-Sheet",title:"01-03 Excel-Object-Model-Cheat-Sheet",description:"Excel Object Model Cheat Sheet",source:"@site/docs/01-cheat-sheets/01-03 Excel-Object-Model-Cheat-Sheet.md",sourceDirName:"01-cheat-sheets",slug:"/cheat-sheets/01-03 Excel-Object-Model-Cheat-Sheet",permalink:"/docs/cheat-sheets/01-03 Excel-Object-Model-Cheat-Sheet",editUrl:"https://github.com/facebook/docusaurus/tree/main/packages/create-docusaurus/templates/shared/docs/01-cheat-sheets/01-03 Excel-Object-Model-Cheat-Sheet.md",tags:[],version:"current",frontMatter:{},sidebar:"tutorialSidebar",previous:{title:"01-02 VBA-Cheat-Sheet",permalink:"/docs/cheat-sheets/01-02 VBA-Cheat-Sheet"},next:{title:"overview of intro, overview",permalink:"/docs/overview-of-intro-overview/02 overview-of-intro-overview"}},s=[{value:"Excel Object Model Cheat Sheet",id:"excel-object-model-cheat-sheet",children:[],level:2}],p={toc:s};function h(e){var t=e.components,r=(0,a.Z)(e,l);return(0,o.kt)("wrapper",(0,n.Z)({},p,r,{components:t,mdxType:"MDXLayout"}),(0,o.kt)("h2",{id:"excel-object-model-cheat-sheet"},"Excel Object Model Cheat Sheet"),(0,o.kt)("p",null,"This cheat sheet will provide a quick glance at the most commonly used objects in Excel and how they are related. It is meant to be a useful check when you know what you want to work with but are not certain how best to get there."),(0,o.kt)("p",null,"TODO: finish this list"),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("p",{parentName:"li"},"Application")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("p",{parentName:"li"},"Workbooks -> Workbook"),(0,o.kt)("ul",{parentName:"li"},(0,o.kt)("li",{parentName:"ul"},"Worksheets -> Worksheet"),(0,o.kt)("li",{parentName:"ul"},"Range -> Range",(0,o.kt)("ul",{parentName:"li"},(0,o.kt)("li",{parentName:"ul"},"Formula"),(0,o.kt)("li",{parentName:"ul"},"Value"),(0,o.kt)("li",{parentName:"ul"},"Address"),(0,o.kt)("li",{parentName:"ul"},"[formatting things]"),(0,o.kt)("li",{parentName:"ul"},"Cells / Rows / Columns"))),(0,o.kt)("li",{parentName:"ul"},"Cells -> Range"),(0,o.kt)("li",{parentName:"ul"},"ChartObjects -> ChartObject",(0,o.kt)("ul",{parentName:"li"},(0,o.kt)("li",{parentName:"ul"},"Chart",(0,o.kt)("ul",{parentName:"li"},(0,o.kt)("li",{parentName:"ul"},"Series"),(0,o.kt)("li",{parentName:"ul"},"Axes -> Axis"),(0,o.kt)("li",{parentName:"ul"},"ChartArea"),(0,o.kt)("li",{parentName:"ul"},"PlotArea"))))),(0,o.kt)("li",{parentName:"ul"},"Shapes -> Shape"),(0,o.kt)("li",{parentName:"ul"},"Names -> Name"),(0,o.kt)("li",{parentName:"ul"},"RefersToRange -> Range")))))}h.isMDXComponent=!0}}]);